using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using Moq;

namespace HtmlToOpenXml.Tests;

/// <summary>
/// Tests parser with various complex input Html.
/// </summary>
[TestFixture]
public class ParserTests : HtmlConverterTestBase
{
    [TestCase("<!--<p>some text</p>-->")]
    [TestCase("<script>document.getElementById('body');</script>")]
    [TestCase("<style>{font-size:2em}</style>")]
    [TestCase("<xml><element><childElement attr='value' /></element></xml>")]
    [TestCase("<button>Save</button>")]
    [TestCase("<input type='search' placeholder='Search' />")]
    [TestCase("<progress>max='100' value='70'>70%</progress>")]
    [TestCase("<select><option>--Please select--</option></select>")]
    [TestCase("<textarea>Placeholder</textarea>")]
    [TestCase("<meter min='200' max='500' value='350'>350 degrees</meter>")]
    [TestCase("<h1><!--empty--></h1>")]
    public void UnsupportedTag_ShouldBeIgnored(string html)
    {
        // the inner html shouldn't be interpreted
        var elements = converter.Parse(html);
        Assert.That(elements, Is.Empty);
    }

    [Test]
    public void Paragraph_WithUnclosedTags_ShouldApplyStyle()
    {
        var elements = converter.Parse("<p>some text in <i>italics <b>,bold and italics</p>");
        Assert.That(elements, Has.Count.EqualTo(1));
        Assert.That(elements[0].InnerText, Is.EqualTo("some text in italics ,bold and italics"));

        var runProperties = elements[0].ChildElements[0].GetFirstChild<RunProperties>();
        Assert.That(runProperties, Is.Null);

        runProperties = elements[0].ChildElements[1].GetFirstChild<RunProperties>();
        Assert.That(runProperties, Is.Not.Null);
        using (Assert.EnterMultipleScope())
        {
            Assert.That(runProperties.HasChild<Italic>(), Is.True);
            Assert.That(runProperties.HasChild<Bold>(), Is.False);
        }

        runProperties = elements[0].ChildElements[2].GetFirstChild<RunProperties>();
        Assert.That(runProperties, Is.Not.Null);
        using (Assert.EnterMultipleScope())
        {
            Assert.That(runProperties.HasChild<Italic>(), Is.True);
            Assert.That(runProperties.HasChild<Bold>(), Is.True);
        }
    }

    [Test]
    public void ConsecutiveParagraph_WithUnclosedTags_ShouldContinueStyle()
    {
        var elements = converter.Parse("<p>First paragraph in semi-<i>italics <p>Second paragraph still italic <b>but also in bold</b></p>");
        Assert.That(elements, Has.Count.EqualTo(2));
        using (Assert.EnterMultipleScope())
        {
            Assert.That(elements[0].Elements<Run>().Count, Is.EqualTo(2));
            Assert.That(elements[1].Elements<Run>().Count, Is.EqualTo(2));
        }

        var runProperties = elements[0].ChildElements[0].GetFirstChild<RunProperties>();
        Assert.That(runProperties, Is.Null);

        runProperties = elements[0].GetLastChild<Run>()!.GetFirstChild<RunProperties>();
        Assert.That(runProperties, Is.Not.Null);
        Assert.That(runProperties.HasChild<Italic>(), Is.True);

        runProperties = elements[1].FirstChild?.GetFirstChild<RunProperties>();
        Assert.That(runProperties, Is.Not.Null);
        Assert.That(runProperties.HasChild<Italic>(), Is.True);
        Assert.That(runProperties.HasChild<Bold>(), Is.False);

        runProperties = elements[1].GetLastChild<Run>()!.GetFirstChild<RunProperties>();
        Assert.That(runProperties, Is.Not.Null);
        Assert.That(runProperties.HasChild<Italic>(), Is.True);
        Assert.That(runProperties.HasChild<Bold>(), Is.True);
    }

    [Test]
    public void ConsecutiveParagraph_WithClosedTags_ShouldNotContinueStyle()
    {
        // this should generate a new paragraph with its own style
        var elements = converter.Parse("<p>First paragraph in <i>italics </i><p>Second paragraph not in italic</p>");
        Assert.That(elements, Has.Count.EqualTo(2));
        Assert.That(elements[0].Elements<Run>().Count, Is.EqualTo(2));
        Assert.That(elements[1].Elements<Run>().Count, Is.EqualTo(1));
        Assert.That(elements[1].FirstChild, Is.TypeOf(typeof(Run)));

        var runProperties = elements[1].FirstChild!.GetFirstChild<RunProperties>();
        Assert.That(runProperties, Is.Null);
    }

    [TestCase("<p>Some\ntext</p>", ExpectedResult = 1)]
    [TestCase("<p>Some <b>bold\n</b>text</p>", ExpectedResult = 3)]
    [TestCase("\t<p>Some <b>bold\n</b>text</p>", ExpectedResult = 3)]
    [TestCase("  <p>Some text</p> ", ExpectedResult = 1)]
    public int Newline_ReturnsRunCount (string html)
    {
        var elements = converter.Parse(html);
        return elements[0].Elements<Run>().Count();
    }

    [TestCase(" < b >bold</b>", ExpectedResult = "< b >bold")]
    [TestCase(" <3", ExpectedResult = "<3")]
    public string EntityNames_ShouldBeTreatedAsSimpleText (string html)
    {
        var elements = converter.Parse(html);
        Assert.That(elements, Has.Count.EqualTo(1));
        Assert.That(elements[0].ChildElements, Has.Count.EqualTo(1));
        Assert.That(elements[0].FirstChild, Is.TypeOf<Run>());
        return elements[0].FirstChild.InnerText;
    }

    [Test(Description = "Provided html is only whitespaces")]
    public void EmptyText_ShouldBeIgnored()
    {
        var elements = converter.Parse("  \n");
        Assert.That(elements, Is.Empty);
    }

    [Test(Description = "Provided mainPart is null")]
    public void ProvidedMainPart_WithNull_ShouldFail()
    {
        Assert.Throws<ArgumentNullException>(() => new HtmlConverter((MainDocumentPart) null!));
    }

    [Test(Description = "Provided mainPart.Document is empty")]
    public void ProvidedMainPartDocument_WithNull_ShouldBeAssigned()
    {
        using var generatedDocument = new MemoryStream();
        using var package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document);
        mainPart = package.MainDocumentPart!;
        mainPart = package.AddMainDocumentPart();

        Assert.That(mainPart.Document, Is.Null);

        var elements = new HtmlConverter(mainPart).Parse("Placeholder");
        Assert.That(elements, Is.Not.Empty);
    }

    [Test(Description = "Provided BaseImageUrl must be an absolute uri")]
    public void ProvidedBaseImageUrl_WithRelativeUrl_ShouldFail()
    {
        Assert.Throws<ArgumentException>(() => new HtmlConverter(mainPart, new IO.DefaultWebRequest {
            BaseImageUrl = new Uri("/path", UriKind.Relative)
        }));
    }

    [Test]
    public async Task LoadHtml_ThenParseFragment_ShouldInsertInCorrectPart()
    {
        var document = await converter.LoadDocumentAsync("""
            <header>Header content</header>
            Body content
        """, TestContext.CurrentContext.CancellationToken);

        var header = document.QuerySelector("header");
        Assert.That(header, Is.Not.Null);
        Assert.That(header.InnerHtml, Is.EqualTo("Header content"));

        await converter.ParseHeader(header);
        header.Remove();

        Assert.That(document.InnerHtml.Trim(), Is.EqualTo("Body content"));

        await converter.ParseBody(document);

        using (Assert.EnterMultipleScope())
        {
            Assert.That(mainPart.HeaderParts.First().Header!.InnerText, Is.EqualTo("Header content"));
            Assert.That(mainPart.Document!.Body!.InnerText, Is.EqualTo("Body content"));
        }
    }

    [Test]
    public void OverrideNodeProcessor_ShouldReplaceDefaultInterpreter()
    {
        var today = DateTime.Today;

        var dateProcessor = new Mock<IHtmlNodeProcessor>();
        dateProcessor.Setup(p => p.CanProcess(It.IsAny<AngleSharp.Dom.INode>()))
            .Returns(false);
        dateProcessor.Setup(p => p.CanProcess(It.Is<AngleSharp.Dom.INode>(
            n => n.NodeName == "H2O-DATE")))
            .Returns(true);

        dateProcessor.Setup(p => p.Interpret(
                It.IsAny<IHtmlNodeContext>(),
                It.Is<AngleSharp.Dom.INode>(n => n.NodeName == "H2O-DATE")))
            .Returns([
                new Run(new Text(today.ToShortDateString()))
            ]);

        converter.NodeProcessors.Add(dateProcessor.Object);

        var elements = converter.Parse("<h2o-date></h2o-date><i>Italics</i>");
        dateProcessor.Verify(p => p.CanProcess(It.IsAny<AngleSharp.Dom.INode>()),
            Times.Exactly(2), "Expected to be called for <h2o-date> and <i>");

        dateProcessor.Verify(p => p.Interpret(
            It.IsAny<IHtmlNodeContext>(),
            It.Is<AngleSharp.Dom.INode>(n => n.NodeName == "H2O-DATE")), Times.Once());

        Assert.That(elements, Has.Count.EqualTo(1));
        Assert.That(elements[0].ChildElements, Has.Count.EqualTo(2));
        Assert.That(elements[0].FirstChild?.InnerText, Is.EqualTo(today.ToShortDateString()));
        Assert.That(elements[0].LastChild?.InnerText, Is.EqualTo("Italics"));
    }
}