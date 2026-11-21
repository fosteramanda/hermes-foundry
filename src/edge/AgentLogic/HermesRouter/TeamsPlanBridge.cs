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
    private readonly Dictionary<string, string> _doneTitles = new(StringComparer.Ordinal);
    private PlanStep? _active;
    private PlanStep? _thinking;
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

    /// <summary>A tool started. Its step joins the Plan under its running ("-ing") name.</summary>
    public async Task StepStartedAsync(string callId, string runningTitle, string doneTitle,
        CancellationToken ct)
    {
        if (Broken) return;
        PlanStep step;
        if (_thinking != null)
        {
            step = _thinking;
            _thinking = null;
            step.Title = Trim(runningTitle);
            step.Version++;
        }
        else
        {
            step = new PlanStep
            {
                Id = $"step-{++_next}",
                Version = 1,
                Title = Trim(runningTitle),
                Status = PlanStepStatuses.Active
            };
        }
        _steps[callId] = step;
        _doneTitles[callId] = Trim(doneTitle);
        _active = step;
        await SendAsync([step], ct);
    }

    /// <summary>A tool finished. Its step is renamed to the past tense and marked done.</summary>
    public async Task StepCompletedAsync(string callId, CancellationToken ct)
    {
        if (Broken || !Started || !_steps.TryGetValue(callId, out var step)) return;
        if (step.Status == PlanStepStatuses.Completed) return;
        MarkDone(callId, step);
        await SendAsync([step], ct);
    }

    public int StepCount => _steps.Count;

    /// <summary>The Plan only ever held its "Thinking" step: no real work ran.</summary>
    public bool OnlyThinking => _thinking != null && _steps.Count == 0;

    private void MarkDone(string callId, PlanStep step)
    {
        step.Status = PlanStepStatuses.Completed;
        if (_doneTitles.TryGetValue(callId, out var done)) step.Title = done;
        step.Version++;
    }

    private async Task<bool> CloseOpenStepsAsync(CancellationToken ct)
    {
        var open = _steps.Where(kv => kv.Value.Status != PlanStepStatuses.Completed).ToList();
        foreach (var (callId, s) in open) MarkDone(callId, s);
        if (open.Count > 0) await _plan.UpdateAsync(open.Select(kv => kv.Value).ToList(), ct);
        return true;
    }

    /// <summary>Completes the Plan with the answer message itself (layout B).</summary>
    public async Task<bool> CompleteAsync(MessageActivityInput answer, CancellationToken ct)
    {
        if (Broken || !Started) return false;
        try
        {
            await CloseOpenStepsAsync(ct);
            if (string.IsNullOrWhiteSpace(answer.Text)) answer.Text = "\u200B";
            await _plan.CompleteAsync(answer, ct);
            return true;
        }
        catch (Exception ex)
        {
            Fail(ex);
            return false;
        }
    }

    /// <summary>
    /// Completes the Plan with an empty completion message, after the answer was sent on its own
    /// (layout A2). Teams refuses a completing message with empty text, so it carries one
    /// zero-width space.
    /// </summary>
    public async Task<bool> CompleteEmptyAsync(CancellationToken ct)
    {
        if (Broken || !Started) return false;
        try
        {
            await CloseOpenStepsAsync(ct);
            await _plan.CompleteAsync(new MessageActivityInput { Text = "\u200B" }, ct);
            return true;
        }
        catch (Exception ex)
        {
            Fail(ex);
            return false;
        }
    }

    /// <summary>Stops the Plan after the answer was sent on its own (layout A1).</summary>
    public async Task<bool> StopAsync(CancellationToken ct)
    {
        if (Broken || !Started) return false;
        try
        {
            await CloseOpenStepsAsync(ct);
            await _plan.StopAsync(ct);
            return true;
        }
        catch (Exception ex)
        {
            Fail(ex);
            return false;
        }
    }

    /// <summary>
    /// The reasoning-only case: no first message and no tool step. Starts the Plan with one
    /// already completed step that carries the why, then completes it, either with the answer
    /// itself (layout b) or, when <paramref name="answer"/> is null, with an empty message sent
    /// after the answer (layout a2).
    /// </summary>
    public async Task<bool> ReasonedOnlyAsync(string why, MessageActivityInput? answer, CancellationToken ct)
    {
        if (Broken) return false;
        try
        {
            var step = new PlanStep
            {
                Id = "why",
                Version = 1,
                // The row name must fit without "..."; the one or two sentences on why go under
                // its arrow as the description.
                Title = "Answered from what I know",
                Description = string.IsNullOrWhiteSpace(why) ? null : why.Trim(),
                Status = PlanStepStatuses.Completed
            };
            await _plan.StartAsync([step], ct);
            Started = true;
            var completion = answer ?? new MessageActivityInput { Text = "\u200B" };
            if (string.IsNullOrWhiteSpace(completion.Text)) completion.Text = "\u200B";
            await _plan.CompleteAsync(completion, ct);
            return true;
        }
        catch (Exception ex)
        {
            Fail(ex);
            return false;
        }
    }

    /// <summary>
    /// Starts the Plan as soon as the first message posts, so the work line reads "Thinking" until
    /// the first step starts. Teams refuses a Plan with no steps ("at least one step is required"),
    /// so it opens with one active step named "Thinking" that the first real step then takes over:
    /// same step, renamed, so "Thinking" never stays behind as a row.
    /// </summary>
    public async Task<bool> StartThinkingAsync(CancellationToken ct)
    {
        if (Broken || Started) return false;
        _thinking = new PlanStep
        {
            Id = $"step-{++_next}",
            Version = 1,
            Title = "Thinking",
            Status = PlanStepStatuses.Active
        };
        await SendAsync([_thinking], ct);
        return !Broken;
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
