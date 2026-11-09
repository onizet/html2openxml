using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.CssInliner.Tests;

/// <summary>
/// Tests parser with various complex Html <c>&lt;style&gt;</c> tags.
/// </summary>
[TestFixture]
public class InlineCssTests
{
    private MemoryStream generatedDocument = default!;
    private WordprocessingDocument package = default!;
    private CssStyleHtmlConverter converter;


    [SetUp]
    public void Init ()
    {
        generatedDocument = new MemoryStream();
        package = WordDocumentFactory.Create(generatedDocument);

        converter = new CssStyleHtmlConverter(package);
    }

    [TearDown]
    public void Close ()
    {
        package?.Dispose();
        generatedDocument?.Dispose();
    }

    [Test]
    public void ParseStyleTag_ShouldBeInlined()
    {
        var elements = converter.Parse("""
            <style>p#para2 { color: red }</style>
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
    }

    [Test]
    public void ParsePriorityStyleTag_ShouldBeOverriden()
    {
        var elements = converter.Parse("""
            <style>p#para2 { color: red !important; }</style>
            <p>Black text</p>
            <p id="para2" style="color:green">Colored text</p>
        """);

        Assert.That(elements, Has.Count.EqualTo(2));
        Assert.That(elements, Has.All.TypeOf<Paragraph>());
        using (Assert.EnterMultipleScope())
        {
            Assert.That(elements[0].GetFirstChild<Run>()?.RunProperties?.Color?.Val?.Value, Is.Null);
            Assert.That(elements[1].GetFirstChild<Run>()?.RunProperties?.Color?.Val?.Value, Is.EqualTo("FF0000"));
        }
    }

    [Test]
    public void ParseLowerPriorityStyleTag_ShouldBeIgnored()
    {
        var elements = converter.Parse("""
            <style>p#para2 { color: red; }</style>
            <p>Black text</p>
            <p id="para2" style="color:green !important">Colored text</p>
        """);

        Assert.That(elements, Has.Count.EqualTo(2));
        Assert.That(elements, Has.All.TypeOf<Paragraph>());
        using (Assert.EnterMultipleScope())
        {
            Assert.That(elements[0].GetFirstChild<Run>()?.RunProperties?.Color?.Val?.Value, Is.Null);
            Assert.That(elements[1].GetFirstChild<Run>()?.RunProperties?.Color?.Val?.Value, Is.EqualTo("008000"));
        }
    }
}