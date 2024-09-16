using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests on <c>body</c> elements.
    /// </summary>
    [TestFixture]
    public class BodyTests : HtmlConverterTestBase
    {
        [TestCase("landscape", ExpectedResult = true)]
        [TestCase("portrait", ExpectedResult = false)]
        public bool PageOrientation_ReturnsLandscapeDimension(string orientation)
        {
            var _ = converter.Parse($@"<body style=""page-orientation:{orientation}""><body>");
            var sectionProperties = mainPart.Document.Body!.GetFirstChild<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Null);
            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            Assert.That(pageSize, Is.Not.Null);
            return pageSize.Width > pageSize.Height;
        }

        [TestCase("portrait", ExpectedResult = true)]
        [TestCase("landscape", ExpectedResult = false)]
        public bool PageOrientation_OverrideExistingLayout_ReturnsLandscapeDimension(string orientation)
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithLandscape.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;
            HtmlConverter converter = new(mainPart);

            var _ = converter.Parse($@"<body style=""page-orientation:{orientation}""><body>");
            var sectionProperties = mainPart.Document.Body!.GetFirstChild<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Null);
            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            Assert.That(pageSize, Is.Not.Null);
            return pageSize.Height > pageSize.Width;
        }

        [TestCase("rtl", ExpectedResult = true)]
        [TestCase("ltr", ExpectedResult = false)]
        [TestCase("", ExpectedResult = null)]
        public bool? WithRtl_ReturnsBidi_DocumentScoped(string dir)
        {
            var elements = converter.Parse($@"<body dir='{dir}'>Lorem</body>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());

            var bidi = mainPart.Document.Body!.GetFirstChild<SectionProperties>()?.GetFirstChild<BiDi>();
            return bidi?.Val?.Value;
        }

        [Test(Description = "Bookmark _GoBack is reserved and must stand after the last edit change")]
        public async Task WithGoBackBookmark_ShouldBeAfterAppendedOutput()
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithCustomStyle.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;

            var goBackBookmark = mainPart.Document.Body!.Elements<Paragraph>()
                .FirstOrDefault(p => p.GetFirstChild<BookmarkStart>()?.Name?.Value == "_GoBack");
            Assert.That(goBackBookmark, Is.Not.Null);

            HtmlConverter converter = new HtmlConverter(mainPart);
            await converter.ParseHtml("<p>Placeholder</p>");

            Assert.That(mainPart.Document.Body!.LastChild, Is.TypeOf<SectionProperties>());
            var paragrahs = mainPart.Document.Body!.Elements<Paragraph>();
            Assert.That(paragrahs.Count(), Is.EqualTo(2));
            Assert.That(paragrahs.Last().GetFirstChild<BookmarkStart>()?.Name?.Value, Is.EqualTo("_GoBack"));
        }
    }
}