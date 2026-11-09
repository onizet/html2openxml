using NUnit.Framework;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Moq;

namespace HtmlToOpenXml.CssInliner.Tests;

/// <summary>
/// CSS Tests parser for stylesheets.
/// </summary>
[TestFixture]
public class StylesheetTests
{
    private MemoryStream generatedDocument = default!;
    private WordprocessingDocument package = default!;
    private CssStyleHtmlConverter converter;
    private Mock<IO.IWebRequest> webRequest;


    [SetUp]
    public void Init ()
    {
        webRequest = new Mock<IO.IWebRequest>();
        webRequest.Setup(x => x.FetchAsync(It.IsAny<Uri>(), It.IsAny<CancellationToken>()))
            .Returns(Task.FromResult<IO.Resource?>(new() {
                Content = new MemoryStream(System.Text.Encoding.ASCII.GetBytes("""
                    p#para2 { color: green }
                """)),
                StatusCode = System.Net.HttpStatusCode.OK
            }));
        webRequest.Setup(x => x.SupportsProtocol(It.Is<string>(s => s == "http")))
            .Returns(true);

        generatedDocument = new MemoryStream();
        package = WordDocumentFactory.Create(generatedDocument);
        converter = new CssStyleHtmlConverter(package, webRequest.Object);
    }

    [TearDown]
    public void Close ()
    {
        package?.Dispose();
        generatedDocument?.Dispose();
    }

    [Test]
    public void ExternalStylesheet_ShouldDownload()
    {
        var elements = converter.Parse("""
            <link href="style1.css" rel="stylesheet" />
            <p>Black text</p>
            <p id="para2">Colored text</p>
        """);

        Assert.That(elements, Has.Count.EqualTo(2));
        Assert.That(elements, Has.All.TypeOf<Paragraph>());
        using (Assert.EnterMultipleScope())
        {
            Assert.That(elements[0].GetFirstChild<Run>()?.RunProperties?.Color?.Val?.Value, Is.Null);
            Assert.That(elements[1].GetFirstChild<Run>()?.RunProperties?.Color?.Val?.Value, Is.EqualTo("008000"));
        }

        webRequest.Verify(s => s.FetchAsync(
            It.Is<Uri>(u => u.AbsolutePath.EndsWith("style1.css")),
            It.IsAny<CancellationToken>()),
        Times.Once);
    }

    [Test]
    public void ExternalStylesheet_ShouldResolvePriorStyleTag()
    {
        var elements = converter.Parse("""
            <style>p#para2 { color: red; }</style>
            <link href="style1.css" rel="stylesheet" />
            <p>Black text</p>
            <p id="para2">Colored text</p>
        """);

        Assert.That(elements, Has.Count.EqualTo(2));
        Assert.That(elements, Has.All.TypeOf<Paragraph>());
        using (Assert.EnterMultipleScope())
        {
            Assert.That(elements[0].GetFirstChild<Run>()?.RunProperties?.Color?.Val?.Value, Is.Null);
            Assert.That(elements[1].GetFirstChild<Run>()?.RunProperties?.Color?.Val?.Value, Is.EqualTo("FF0000"));
        }

        webRequest.Verify(s => s.FetchAsync(
            It.Is<Uri>(u => u.AbsolutePath.EndsWith("style1.css")),
            It.IsAny<CancellationToken>()),
        Times.Once);
    }
}