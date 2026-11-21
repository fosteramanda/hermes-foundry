namespace HelloWorldA365.AgentLogic;

using Microsoft.Agents.Builder;
using Microsoft.Agents.Core.Models;

/// <summary>
/// The Teams typing indicator for one turn: sent as soon as the message arrives, then again every
/// 3 seconds (Teams drops it after about 5), until the first thing posts: the answer when Hermes
/// needs no tools, or the first message when it does. Teams shows typing only in one-on-one and
/// small group chats; it is sent everywhere and simply not drawn elsewhere.
/// </summary>
public sealed class TypingLoop
{
    public const string StateKey = "Hermes.Typing";
    private static readonly TimeSpan Interval = TimeSpan.FromSeconds(3);

    private readonly CancellationTokenSource _stop;
    private readonly Task _loop;

    private TypingLoop(ITurnContext turnContext, CancellationToken turnToken)
    {
        _stop = CancellationTokenSource.CreateLinkedTokenSource(turnToken);
        _loop = Task.Run(() => RunAsync(turnContext, _stop.Token));
    }

    public static TypingLoop Start(ITurnContext turnContext, CancellationToken turnToken)
    {
        var loop = new TypingLoop(turnContext, turnToken);
        turnContext.Services.Set(StateKey, loop);
        return loop;
    }

    /// <summary>Stops the loop for this turn, if one is running. Safe to call more than once.</summary>
    public static void StopFor(ITurnContext turnContext)
    {
        try { turnContext?.Services.Get<TypingLoop>(StateKey)?.Stop(); }
        catch { /* the indicator is decoration; never fail a reply over it */ }
    }

    public void Stop()
    {
        if (!_stop.IsCancellationRequested) _stop.Cancel();
    }

    private static async Task RunAsync(ITurnContext turnContext, CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            try
            {
                await turnContext.SendActivityAsync(new Activity { Type = ActivityTypes.Typing }, ct);
            }
            catch (OperationCanceledException) { return; }
            catch { /* a dropped typing activity is not worth failing over */ }
            try { await Task.Delay(Interval, ct); }
            catch (OperationCanceledException) { return; }
        }
    }
}
