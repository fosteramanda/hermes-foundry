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

        Assert.Equal("Searching for biggest news", running);
        Assert.Equal("Searched for biggest news", done);
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
        Assert.Equal("Searching for single", running);
    }

    [Fact]
    public void RepeatedWorkGetsADifferentNameEachTime()
    {
        var names = new StepNamer();
        var page = Args(("url", "https://www.reuters.com/world/"));

        var first = names.Name("fetch_url", page).Running;
        var second = names.Name("browse_url", page).Running;
        var third = names.Name("fetch_url", Args(("url", "https://apnews.com/article/summit-peace-declaration-123")));

        Assert.Equal("Reading reuters.com", first);
        Assert.Equal("Reading reuters.com again", second);
        Assert.Equal("Reading apnews.com", third.Running);
        Assert.Equal("Read apnews.com", third.Done);
    }

    [Theory]
    [InlineData("mcp__m365__start_task")]
    [InlineData("skill_view")]
    public void BookkeepingToolsAreNotSteps(string tool) => Assert.True(StepNamer.IsHidden(tool));
}
