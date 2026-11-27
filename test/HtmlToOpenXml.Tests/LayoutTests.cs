using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.Tests;

/// <summary>
/// Tests parser with the document layout helpers.
/// </summary>
[TestFixture]
public class LayoutTests : HtmlConverterTestBase
{
    [Test]
    public async Task AddPageBreak_ShouldCreateBreakAfterTOC()
    {
        converter.AddPageBreak();

        AssertThatOpenXmlDocumentIsValid();
        var run = mainPart.Document!.Body!.GetFirstChild<Paragraph>()?.GetFirstChild<Run>();
        Assert.That(run, Is.Not.Null);
        Assert.That(run?.HasChild<Break>(), Is.True);
        Assert.That(run?.GetFirstChild<Break>()?.Type?.Value, Is.EqualTo(BreakValues.Page));

        var sectionProperties = mainPart.Document!.Body!.GetLastChild<SectionProperties>();
        Assert.That(sectionProperties, Is.Not.Null);
        Assert.That(sectionProperties.IsAfter(run), Is.True, "SectionProperties must remains the last child");
    }

    [Test]
    public async Task AddPageWithDifferentOrientation_ShouldCreateSectionBreak()
    {
        await converter.ParseBody("Page Portrait");

        converter.AddPage(PageOrientation.Landscape);
        await converter.ParseBody("Page Landscape");

        AssertThatOpenXmlDocumentIsValid();

        var firstPage = mainPart.Document!.Body!.GetFirstChild<Paragraph>();
        Assert.That(firstPage?.InnerText, Is.EqualTo("Page Portrait"));
        var paraBreakSection = firstPage.NextSibling();
        Assert.That((paraBreakSection as Paragraph)?.GetFirstChild<ParagraphProperties>()?.SectionProperties, Is.Not.Null,
            "The next paragraph must enclose a SectionProperties");
        var pageSize = paraBreakSection!.GetFirstChild<ParagraphProperties>()?.SectionProperties!.GetFirstChild<PageSize>();
        Assert.That(pageSize, Is.Not.Null);
        Assert.That(pageSize?.IsLandscape(), Is.False,
            "The first page layout is ending with this Portrait section properties");
        Assert.That(paraBreakSection.GetFirstChild<Run>()?.HasChild<LastRenderedPageBreak>(), Is.True);

        var secondPage = paraBreakSection.NextSibling();
        Assert.That((secondPage as Paragraph)?.InnerText, Is.EqualTo("Page Landscape"));

        var lastChild = mainPart.Document!.Body!.LastChild;
        Assert.That(lastChild, Is.TypeOf<SectionProperties>());
        Assert.That(lastChild.GetFirstChild<PageSize>()?.IsLandscape(), Is.True);
    }

    [Test]
    public void AddPageOrientation_ToExistingLandscapeTemplate_ShouldAddSectionBreak()
    {
        using var generatedDocument = new MemoryStream();
        using (var buffer = ResourceHelper.GetStream("Resources.DocWithLandscape.docx"))
            buffer.CopyTo(generatedDocument);

        generatedDocument.Position = 0L;
        using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
        HtmlConverter converter = new(package);

        converter.AddPage(PageOrientation.Portrait);
        AssertThatOpenXmlDocumentIsValid();

        var body = package.MainDocumentPart!.Document!.Body!;
        var paraBreakSection = body.Elements<Paragraph>().FirstOrDefault(p => p.ParagraphProperties?.SectionProperties != null);
        Assert.That(paraBreakSection, Is.Not.Null, "The paragraph must enclose a SectionProperties");
        var pageSize = paraBreakSection!.GetFirstChild<ParagraphProperties>()?.SectionProperties!.GetFirstChild<PageSize>();
        Assert.That(pageSize, Is.Not.Null);
        Assert.That(pageSize?.IsLandscape(), Is.True,
            "The first page layout remains Landscape");

        var sectionProperties = body.GetLastChild<SectionProperties>();
        Assert.That(sectionProperties, Is.Not.Null);
        pageSize = sectionProperties.GetFirstChild<PageSize>();
        Assert.That(pageSize, Is.Not.Null);
        Assert.That(pageSize.IsLandscape(), Is.False);
    }

    [Test]
    public void AddPageLandscape_ToExistingLandscapeTemplate_ShouldPreserveOrientation_But_AddPageBreak()
    {
        using var generatedDocument = new MemoryStream();
        using (var buffer = ResourceHelper.GetStream("Resources.DocWithLandscape.docx"))
            buffer.CopyTo(generatedDocument);

        generatedDocument.Position = 0L;
        using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
        HtmlConverter converter = new(package);

        converter.AddPage(PageOrientation.Landscape);
        AssertThatOpenXmlDocumentIsValid();

        var body = package.MainDocumentPart!.Document!.Body!;
        var paraBreakSection = body.Elements<Paragraph>().FirstOrDefault(p => p.ParagraphProperties?.SectionProperties != null);
        Assert.That(paraBreakSection, Is.Null, "No paragraph must enclose a SectionProperties as this is the same page size");

        var paraBreakPage = body.GetFirstChild<Paragraph>()?.GetLastChild<Run>()?.GetFirstChild<Break>();
        Assert.That(paraBreakPage?.Type?.Value, Is.EqualTo(BreakValues.Page));

        var sectionProperties = body.GetLastChild<SectionProperties>();
        Assert.That(sectionProperties, Is.Not.Null);
        var pageSize = sectionProperties.GetFirstChild<PageSize>();
        Assert.That(pageSize, Is.Not.Null);
        Assert.That(pageSize.IsLandscape(), Is.True);
    }
}