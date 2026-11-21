namespace HelloWorldA365Agent.Tests;

using HelloWorldA365.AgentLogic.HermesRouter;
using Xunit;

public sealed class StepNamerTests
{
    private static Dictionary<string, string> Args(params (string k, string v)[] pairs) =>
        pairs.ToDictionary(p => p.k, p => p.v);

    [Fact]
    public void RunningAndDoneNamesAreFullSpecificAndChangeTense()
    {
        var names = new StepNamer();

        var (running, done) = names.Name("mcp__m365__web_search", Args(("query", "biggest news this week")));

        Assert.Equal("Searching the web for \"biggest news this week\"", running);
        Assert.Equal("Searched the web for \"biggest news this week\"", done);
    }

    [Fact]
    public void LongNamesAreShortenedAtAWordWithoutAnEllipsis()
    {
        var names = new StepNamer();

        var (running, done) = names.Name("web_search",
            Args(("query", "what is the single biggest world news story of this week according to major outlets")));

        Assert.True(running.Length <= StepNamer.MaxLength, running);
        Assert.True(done.Length <= StepNamer.MaxLength, done);
        Assert.DoesNotContain("...", running);
        Assert.EndsWith("\"", running);
    }

    [Fact]
    public void RepeatedWorkGetsADifferentNameEachTime()
    {
        var names = new StepNamer();
        var page = Args(("url", "https://www.reuters.com/world/"));

        var first = names.Name("fetch_url", page).Running;
        var second = names.Name("browse_url", page).Running;
        var third = names.Name("fetch_url", Args(("url", "https://apnews.com/article/summit-peace-declaration-123")));

        Assert.Equal("Reading reuters.com: world", first);
        Assert.NotEqual(first, second);
        Assert.Equal("Reading apnews.com: summit peace declaration", third.Running);
        Assert.Equal("Read apnews.com: summit peace declaration", third.Done);
    }

    [Theory]
    [InlineData("mcp__m365__start_task")]
    [InlineData("skill_view")]
    public void BookkeepingToolsAreNotSteps(string tool) => Assert.True(StepNamer.IsHidden(tool));
}
