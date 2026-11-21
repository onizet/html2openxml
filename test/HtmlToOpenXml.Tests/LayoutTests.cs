using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests;

/// <summary>
/// Tests parser with the document layout helpers.
/// </summary>
[TestFixture]
public class DocumentLayoutTests : HtmlConverterTestBase
{
    [Test]
    public async Task SetCoverPage_ShouldAddAsFirst()
    {
        await converter.ParseBody("Body placeholder");
        await converter.SetCoverPage("<h1>Cover Page</h1>");

        AssertThatOpenXmlDocumentIsValid();

        var body = mainPart.Document!.Body!;
        var cover = body.GetFirstChild<SdtBlock>();
        var paragraph = body.GetFirstChild<Paragraph>();

        using (Assert.EnterMultipleScope())
        {
            Assert.That(cover, Is.Not.Null, "Cover SdtBlock is first");
            Assert.That(paragraph?.InnerText, Is.EqualTo("Body placeholder"));
        }

        Assert.That(body.ChildElements.IndexOf(cover),
            Is.LessThan(body.ChildElements.IndexOf(paragraph)));
    }

    [Test]
    public async Task SetCoverPage_ShouldReplaceExisting()
    {
        await converter.SetCoverPage("<h1>Cover Page</h1>");

        await converter.SetCoverPage("<h1>Cover Page 2</h1>");

        AssertThatOpenXmlDocumentIsValid();

        var coverSdtBlock = converter.GetCoverPage();
        Assert.That(coverSdtBlock, Is.Not.Null);
        Assert.That(coverSdtBlock.SdtContentBlock?.InnerText, Is.EqualTo("Cover Page 2"));
    }

    [Test]
    public async Task SetTOC_ShouldCreateNativeTOC()
    {
        const string legend = "Table de matières";
        await converter.ParseBody("""
            <h1>Heading I</h1>
                <h2>Heading I.II</h2>
            <h1>Heading II</h1>
        """);
        converter.SetTableOfContents(maxLevel: 1, caption: legend);

        AssertThatOpenXmlDocumentIsValid();

        var tocSdtBlock = converter.GetTableOfContents();
        Assert.That(tocSdtBlock, Is.Not.Null);
        Assert.That(tocSdtBlock.SdtContentBlock, Is.Not.Null);
        var paragraphs = tocSdtBlock.SdtContentBlock.Elements<Paragraph>().ToList();
        Assert.That(paragraphs.Count, Is.EqualTo(4), "TOC Heading + 2x headings lvl1 + fieldChar:end");

        Assert.That(paragraphs[0].InnerText, Is.EqualTo(legend));
        Assert.That(paragraphs[1].ChildElements.OfType<Run>().Select(r => r.GetFirstChild<FieldCode>()?.InnerText),
            Has.One.EqualTo(@" TOC \o ""1-1"" \h \z \u "));

        Assert.That(tocSdtBlock.SdtProperties?
            .Descendants<DocPartGallery>().Single().Val?.Value, Is.EqualTo("Table of Contents"),
            "This marker ensures Word recognize the content control as a TOC");
        Assert.That(tocSdtBlock.SdtProperties?.Descendants<DocPartUnique>().Single().Val?.Value, Is.True);

        var hyperlinks = tocSdtBlock.SdtContentBlock.Descendants<Hyperlink>().ToList();
        Assert.That(hyperlinks, Has.Count.EqualTo(2));
        Assert.That(hyperlinks.All(h => !string.IsNullOrWhiteSpace(h.Anchor?.Value)), Is.True);

        foreach (var hyperlink in hyperlinks)
        {
            var anchor = hyperlink.Anchor?.Value;
            var pageRef = hyperlink.Descendants<FieldCode>().SingleOrDefault();
            Assert.That(pageRef, Is.Not.Null);
            Assert.That(pageRef.InnerText, Does.StartWith(" PAGEREF "));
            Assert.That(pageRef.InnerText, Does.Contain(anchor));
        }

        var bookmarks = new HashSet<string>(converter.mainPart.Document!.Body!.Descendants<BookmarkStart>()
            .Select(b => b.Name?.Value)!);

        foreach (var hyperlink in hyperlinks)
        {
            Assert.That(bookmarks, Does.Contain(hyperlink.Anchor?.Value));
        }

        Assert.That(tocSdtBlock.SdtContentBlock
            .Descendants<FieldChar>()
            .Count(f => f.FieldCharType?.Value == FieldCharValues.Begin),
        Is.EqualTo(tocSdtBlock.SdtContentBlock
            .Descendants<FieldChar>()
            .Count(f => f.FieldCharType?.Value == FieldCharValues.End)
        ), "Every field begin marker should have a matching field end marker.");
    }

    [Test]
    public async Task AddPageBreakAfter_ShouldCreateBreakAfterTOC()
    {
        var toc = converter.SetTableOfContents();
        converter.AddPageBreakAfter(toc);

        AssertThatOpenXmlDocumentIsValid();
        var run = mainPart.Document!.Body!.GetFirstChild<Paragraph>()?.GetFirstChild<Run>();
        Assert.That(run, Is.Not.Null);
        Assert.That(toc.IsBefore(run), Is.True);
        Assert.That(run?.HasChild<Break>(), Is.True);
        Assert.That(run?.GetFirstChild<Break>()?.Type?.Value, Is.EqualTo(BreakValues.Page));
    }

    [Test]
    public async Task AddPageBreakBefore_ShouldCreateBreakBeforeTOC()
    {
        var toc = converter.SetTableOfContents();
        converter.AddPageBreakBefore(toc);

        AssertThatOpenXmlDocumentIsValid();
        var run = mainPart.Document!.Body!.GetFirstChild<Paragraph>()?.GetFirstChild<Run>();
        Assert.That(run, Is.Not.Null);
        Assert.That(toc.IsAfter(run), Is.True);
        Assert.That(run?.HasChild<Break>(), Is.True);
        Assert.That(run?.GetFirstChild<Break>()?.Type?.Value, Is.EqualTo(BreakValues.Page));
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
}