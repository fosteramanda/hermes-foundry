namespace HelloWorldA365.AgentLogic.HermesRouter;

#pragma warning disable ExperimentalTeamsPlan // the Teams SDK plan API is a documented public preview

using System.Net.Http.Headers;
using System.Reflection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Clients;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Apps.Schema.Entities;
using CoreConversationClient = Microsoft.Teams.Core.ConversationClient;
using CoreUserTokenClient = Microsoft.Teams.Core.UserTokenClient;
using TeamsAgenticIdentity = Microsoft.Teams.Core.Schema.AgenticIdentity;

/// <summary>
/// Teams' native chain of thought ("Plans and steps for Teams agents") for one Hermes turn.
///
/// Every Plan call goes through the Teams SDK's own <see cref="Plan"/> class; nothing here builds a
/// planInfo entity. The SDK's API client is normally created by its own hosting, which this edge
/// does not use (it is a Microsoft 365 Agents SDK app), so the client is constructed directly and
/// authenticated with the agent user's own token for the Teams service, exchanged from this turn.
///
/// The Plan is anchored to Hermes's own first message ("Working on your request..."), so the
/// working line sits under the agent's message, never under the person's. Steps join one at a
/// time as tools start; the answer completes the Plan. A turn that runs no tool sends no Plan.
/// Any refusal from Teams marks the bridge broken and the caller falls back to the progress
/// message it already had.
/// </summary>
internal sealed class TeamsPlanBridge
{
    public const string TeamsScope = "https://botapi.skype.com/.default";
    public const string StateKey = "Hermes.TeamsPlan";

    private readonly Plan _plan;
    private readonly Dictionary<string, PlanStep> _steps = new(StringComparer.Ordinal);
    private PlanStep? _active;
    private int _next;

    public bool Started { get; private set; }
    public bool Broken { get; private set; }
    public string? LastError { get; private set; }

    private TeamsPlanBridge(Plan plan) => _plan = plan;

    /// <summary>On unless HERMES_TEAMS_PLAN is "false".</summary>
    public static bool Enabled =>
        !string.Equals(Environment.GetEnvironmentVariable("HERMES_TEAMS_PLAN"), "false",
            StringComparison.OrdinalIgnoreCase);

    public static TeamsPlanBridge Create(Uri serviceUrl, string conversationId, string anchorActivityId,
        string agentUserToken)
    {
        ArgumentNullException.ThrowIfNull(serviceUrl);
        ArgumentException.ThrowIfNullOrWhiteSpace(conversationId);
        ArgumentException.ThrowIfNullOrWhiteSpace(anchorActivityId);
        ArgumentException.ThrowIfNullOrWhiteSpace(agentUserToken);

        var http = new HttpClient(new BearerHandler(agentUserToken) { InnerHandler = new HttpClientHandler() })
        {
            Timeout = TimeSpan.FromSeconds(20)
        };
        var conversations = new CoreConversationClient(http, NullLogger<CoreConversationClient>.Instance);
        var userTokens = new CoreUserTokenClient(http, new ConfigurationBuilder().Build(),
            NullLogger<CoreUserTokenClient>.Instance);
        return new TeamsPlanBridge(new Plan(CreateApiClient(serviceUrl, http, conversations, userTokens),
            conversationId, anchorActivityId));
    }

    internal static ApiClient CreateApiClient(Uri serviceUrl, HttpClient http,
        CoreConversationClient conversations, CoreUserTokenClient userTokens)
    {
        var ctor = typeof(ApiClient).GetConstructor(
            BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public, null,
            [typeof(Uri), typeof(HttpClient), typeof(CoreConversationClient), typeof(CoreUserTokenClient),
             typeof(Microsoft.Extensions.Logging.ILogger), typeof(TeamsAgenticIdentity)], null)
            ?? throw new MissingMethodException(
                "The Teams SDK ApiClient(serviceUrl, httpClient, conversationClient, userTokenClient, logger, agenticIdentity) constructor was not found; check the Microsoft.Teams.Apps version.");
        return (ApiClient)ctor.Invoke([serviceUrl, http, conversations, userTokens, null, null]);
    }

    /// <summary>A tool started. Its step joins the Plan; the step before it is marked done.</summary>
    public async Task StepStartedAsync(string callId, string title, CancellationToken ct)
    {
        if (Broken) return;
        var changed = new List<PlanStep>();
        if (_active != null && _active.Status != PlanStepStatuses.Completed)
        {
            _active.Status = PlanStepStatuses.Completed;
            _active.Version++;
            changed.Add(_active);
        }
        var step = new PlanStep
        {
            Id = $"step-{++_next}",
            Version = 1,
            Title = Trim(title),
            Status = PlanStepStatuses.Active
        };
        _steps[callId] = step;
        _active = step;
        changed.Add(step);
        await SendAsync(changed, ct);
    }

    /// <summary>A tool finished. Its step is marked done.</summary>
    public async Task StepCompletedAsync(string callId, CancellationToken ct)
    {
        if (Broken || !Started || !_steps.TryGetValue(callId, out var step)) return;
        if (step.Status == PlanStepStatuses.Completed) return;
        step.Status = PlanStepStatuses.Completed;
        step.Version++;
        await SendAsync([step], ct);
    }

    /// <summary>Completes the Plan with the answer. Returns false if Teams refused it.</summary>
    public async Task<bool> CompleteAsync(string answer, CancellationToken ct)
    {
        if (Broken || !Started) return false;
        try
        {
            var open = _steps.Values.Where(s => s.Status != PlanStepStatuses.Completed).ToList();
            foreach (var s in open)
            {
                s.Status = PlanStepStatuses.Completed;
                s.Version++;
            }
            if (open.Count > 0) await _plan.UpdateAsync(open, ct);
            // Teams refuses a completing message with empty text.
            var text = string.IsNullOrWhiteSpace(answer) ? "\u200B" : answer;
            await _plan.CompleteAsync(new MessageActivityInput { Text = text }, ct);
            return true;
        }
        catch (Exception ex)
        {
            Fail(ex);
            return false;
        }
    }

    private async Task SendAsync(List<PlanStep> steps, CancellationToken ct)
    {
        try
        {
            if (!Started)
            {
                // Start carries only what has begun; a completed step from before the start is not sent.
                await _plan.StartAsync(steps.Where(s => s.Status != PlanStepStatuses.Completed).ToList(), ct);
                Started = true;
            }
            else
            {
                await _plan.UpdateAsync(steps, ct);
            }
        }
        catch (Exception ex)
        {
            Fail(ex);
        }
    }

    private void Fail(Exception ex)
    {
        Broken = true;
        LastError = $"{ex.GetType().Name}: {Shorten(ex.Message)}";
    }

    private static string Trim(string title)
    {
        var t = string.IsNullOrWhiteSpace(title) ? "Working" : title.Trim();
        return t.Length <= 120 ? t : t[..117] + "...";
    }

    private static string Shorten(string message)
    {
        var m = (message ?? "").Replace('\n', ' ').Replace('\r', ' ');
        return m.Length <= 300 ? m : m[..300];
    }

    private sealed class BearerHandler(string token) : DelegatingHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            return base.SendAsync(request, cancellationToken);
        }
    }
}
