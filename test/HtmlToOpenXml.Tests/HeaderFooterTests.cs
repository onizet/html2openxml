using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests on <c>ParseHeader</c> and <c>ParseFooter</c> methods.
    /// </summary>
    [TestFixture]
    public class HeaderFooterTests : HtmlConverterTestBase
    {
        [Test]
        public async Task Header_ReturnsHeaderPartLinkedToBody()
        {
            await converter.ParseHeader("<p>Header content</p>", HeaderFooterValues.First);

            var headerPart = mainPart.HeaderParts?.FirstOrDefault();
            Assert.That(headerPart, Is.Not.Null);
            Assert.That(headerPart.Header, Is.Not.Null);
            var p = headerPart.Header.Elements<Paragraph>();
            Assert.That(p, Is.Not.Empty);
            Assert.That(p.Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value), 
                Has.All.EqualTo(converter.HtmlStyles.DefaultStyles.HeaderStyle));

            var sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Empty);
            Assert.That(sectionProperties.SelectMany(s => s.Elements<HeaderReference>())
                .Any(r => r.Type?.Value == HeaderFooterValues.First), Is.True);
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test]
        public async Task Footer_ReturnsFooterPartLinkedToBody()
        {
            await converter.ParseFooter("<p>Footer content</p>");

            var footerPart = mainPart.FooterParts?.FirstOrDefault();
            Assert.That(footerPart, Is.Not.Null);
            Assert.That(footerPart.Footer, Is.Not.Null);

            var sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Empty);
            Assert.That(sectionProperties.Any(s => s.HasChild<FooterReference>()), Is.True);
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test(Description = "Overwrite existing Default header")]
        public async Task WithExistingHeader_Default_ReturnsOverridenHeaderPart()
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithImgHeaderFooter.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;

            var sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Empty);
            var headerRefs = sectionProperties.SelectMany(s => s.Elements<HeaderReference>());
            Assert.Multiple(() =>
            {
                Assert.That(headerRefs.Count(), Is.EqualTo(1));
                Assert.That(headerRefs.Count(r => r.Type?.Value == HeaderFooterValues.Default), Is.EqualTo(1), "Default header exist");
            });

            HtmlConverter converter = new(mainPart);
            await converter.ParseHeader("Header content");

            sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Empty);
            Assert.That(sectionProperties.SelectMany(s => s.Elements<HeaderReference>())
                .Count(r => r.Type?.Value == HeaderFooterValues.Default), Is.EqualTo(1));
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test(Description = "Create additional header for even pages")]
        public async Task WithExistingHeader_Even_ReturnsAnotherHeaderPart()
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithImgHeaderFooter.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;

            var sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Empty);
            var headerRefs = sectionProperties.SelectMany(s => s.Elements<HeaderReference>());
            Assert.Multiple(() =>
            {
                Assert.That(headerRefs.Count(r => r.Type?.Value == HeaderFooterValues.Default), Is.EqualTo(1), "Default header exist");
                Assert.That(headerRefs.Count(r => r.Type?.Value == HeaderFooterValues.Even), Is.Zero, "No event header has been yet defined");
            });

            HtmlConverter converter = new(mainPart);
            await converter.ParseHeader("Header even content", HeaderFooterValues.Even);

            sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Empty);
            Assert.That(sectionProperties.Count(s => s.HasChild<HeaderReference>()), Is.EqualTo(1));
            headerRefs = sectionProperties.SelectMany(s => s.Elements<HeaderReference>());
            Assert.Multiple(() =>
            {
                Assert.That(headerRefs.Count(r => r.Type?.Value == HeaderFooterValues.Default), Is.EqualTo(1));
                Assert.That(headerRefs.Count(r => r.Type?.Value == HeaderFooterValues.Even), Is.EqualTo(1));
            });
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test]
        public async Task Header_ReturnsStyleParagraphs()
        {
            await converter.ParseHeader(@"
                <header>
                <p>Placeholder
                    <nav>
                        <ul><li>Home</li><li>Contact</li></ul>
                    </nav>
                </p>
                </header>
            ");

            var header = mainPart.HeaderParts.FirstOrDefault()?.Header;
            Assert.That(header, Is.Not.Null);
            var paragraphs = header.Elements<Paragraph>();
            Assert.That(paragraphs.Count(), Is.EqualTo(3));
            Assert.That(paragraphs.First().ParagraphProperties?.ParagraphStyleId?.Val?.Value,
                Is.EqualTo(converter.HtmlStyles.DefaultStyles.HeaderStyle));
            Assert.That(paragraphs.Skip(1).Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value),
                Has.All.EqualTo(converter.HtmlStyles.DefaultStyles.ListParagraphStyle));
        }

        [Test]
        public async Task Footer_ReturnsStyleParagraphs()
        {
            await converter.ParseFooter(@"
                <footer>
                <p>
                    <a rel=""license"" href=""https://creativecommons.org/licenses/by/4.0/"">Copyrighted but you can use what's here as long as you credit me</a>
                    <small>&copy; Copyright 2058, Company Inc.</small>
                </p>
                </footer>
            ");

            var footer = mainPart.FooterParts.FirstOrDefault()?.Footer;
            Assert.That(footer, Is.Not.Null);
            var paragraphs = footer.Elements<Paragraph>();
            Assert.That(paragraphs.Count(), Is.EqualTo(2));
            Assert.That(paragraphs.Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value),
                Has.All.EqualTo(converter.HtmlStyles.DefaultStyles.FooterStyle));
        }
    }
}