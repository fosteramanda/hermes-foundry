namespace HelloWorldA365Agent.Tests;

using HelloWorldA365.AgentLogic.HermesRouter;
using Xunit;

public sealed class OfficeLinkFormatterTests
{
    private const string WordUrl =
        "https://contoso-my.sharepoint.com/personal/hermes_contoso_com/_layouts/15/Doc.aspx"
        + "?sourcedoc=%7B1A2B%7D&file=Q3%20Planning%20Notes.docx&action=default&mobileredirect=true";

    [Fact]
    public void BareWordAddressBecomesTitledLinkAndKeepsSentencePunctuation()
    {
        var result = OfficeLinkFormatter.Format($"Done: {WordUrl}.");

        Assert.Equal($"Done: [Q3 Planning Notes]({WordUrl}).", result);
    }

    [Theory]
    [InlineData("https://contoso-my.sharepoint.com/personal/h/Documents/Budget.xlsx", "Budget")]
    [InlineData("https://contoso-my.sharepoint.com/personal/h/Documents/Kickoff%20deck.pptx", "Kickoff deck")]
    [InlineData("https://contoso.sharepoint.com/:x:/g/personal/h/EabcXYZ", "Excel workbook")]
    public void ExcelAndPowerPointAddressesAreTitled(string url, string title)
    {
        Assert.Equal($"See [{title}]({url})", OfficeLinkFormatter.Format($"See <{url}>"));
    }

    [Fact]
    public void GenericLinkTextIsReplacedButAWrittenTitleIsKept()
    {
        Assert.Equal($"[Q3 Planning Notes]({WordUrl})",
            OfficeLinkFormatter.Format($"[here]({WordUrl})"));
        Assert.Equal($"[Planning notes for Thursday]({WordUrl})",
            OfficeLinkFormatter.Format($"[Planning notes for Thursday]({WordUrl})"));
    }

    [Fact]
    public void OtherAddressesAndHtmlAnchorsAreLeftAlone()
    {
        var text = "Read https://learn.microsoft.com/teams and "
            + $"<a href=\"{WordUrl}\">notes</a>, or https://contoso-my.sharepoint.com/personal/h/Documents/report.pdf";

        Assert.Equal(text, OfficeLinkFormatter.Format(text));
    }
}
