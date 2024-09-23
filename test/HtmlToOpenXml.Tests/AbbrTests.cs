using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests acronym, abbreviation and blockquotes.
    /// </summary>
    [TestFixture]
    public class AbbrTests : HtmlConverterTestBase
    {
        [TestCase(@"<dfn title='National Aeronautics and Space Administration'>NASA</dfn>")]
        [TestCase(@"<abbr title='National Aeronautics and Space Administration'>NASA</abbr>")]
        [TestCase(@"<acronym title='National Aeronautics and Space Administration'>NASA</acronym>")]
        [TestCase(@"<acronym title='www.nasa.gov'>NASA</acronym>")]
        public void WithTitle_ReturnsFootnote(string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.Multiple(() => {
                Assert.That(elements[0], Is.TypeOf(typeof(Paragraph)));
                Assert.That(elements[0].HasChild<Run>(), Is.True);
                Assert.That(elements[0].InnerText, Is.EqualTo("NASA"));
            });

            var noteRef = elements[0].GetLastChild<Run>()?.GetFirstChild<FootnoteReference>();
            Assert.That(noteRef, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(noteRef.Id?.HasValue, Is.EqualTo(true));
                Assert.That(mainPart.FootnotesPart, Is.Not.Null);
            });

            Assert.That(mainPart.FootnotesPart.HyperlinkRelationships.Count(), Is.EqualTo(0));

            var fnotes = mainPart.FootnotesPart.Footnotes.Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == noteRef.Id.Value);
            Assert.That(fnotes, Is.Not.Null);
        }

        [TestCase(@"<abbr title='https://en.wikipedia.org/wiki/N A S A '>NASA</abbr>", "https://en.wikipedia.org/wiki/N%20A%20S%20A")]
        [TestCase(@"<abbr title='file://C:\temp\NASA.html'>NASA</abbr>", @"file:///C:/temp/NASA.html")]
        [TestCase(@"<abbr title='\\server01\share\NASA.html'>NASA</abbr>", "file://server01/share/NASA.html")]
        [TestCase(@"<abbr title='ftp://server01/share/NASA.html'>NASA</abbr>", "ftp://server01/share/NASA.html")]
        [TestCase(@"<blockquote cite='https://en.wikipedia.org/wiki/NASA'>NASA</blockquote>", "https://en.wikipedia.org/wiki/NASA")]
        public void WithLink_ReturnsFootnote_WithHyperlink(string html, string expectedUri)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.Multiple(() => {
                Assert.That(elements[0], Is.TypeOf(typeof(Paragraph)));
                Assert.That(elements[0].HasChild<Run>(), Is.True);
                Assert.That(elements[0].InnerText, Is.EqualTo("NASA"));
            });

            var noteRef = elements[0].GetLastChild<Run>()?.GetFirstChild<FootnoteReference>();
            Assert.That(noteRef, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(noteRef.Id?.HasValue, Is.EqualTo(true));
                Assert.That(mainPart.FootnotesPart, Is.Not.Null);
            });

            var fnotes = mainPart.FootnotesPart.Footnotes.Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == noteRef.Id.Value);
            Assert.That(fnotes, Is.Not.Null);

            var link = fnotes.FirstChild?.GetFirstChild<Hyperlink>();
            Assert.That(link, Is.Not.Null);

            var extLink = mainPart.FootnotesPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == link.Id);
            Assert.That(extLink, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(extLink.IsExternal, Is.EqualTo(true));
                Assert.That(extLink.Uri.AbsoluteUri, Is.EqualTo(expectedUri));
            });
        }

        [Test]
        public void WithPositionToDocumentEnd_ReturnsEndnote()
        {
            converter.AcronymPosition = AcronymPosition.DocumentEnd;
            var elements = converter.Parse(@"<acronym title='www.nasa.gov'>NASA</acronym>");

            var noteRef = elements[0].GetLastChild<Run>()?.GetFirstChild<EndnoteReference>();
            Assert.That(noteRef, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(noteRef.Id?.HasValue, Is.EqualTo(true));
                Assert.That(mainPart.EndnotesPart, Is.Not.Null);
            });

            var fnotes = mainPart.EndnotesPart.Endnotes.Elements<Endnote>().FirstOrDefault(f => f.Id?.Value == noteRef.Id.Value);
            Assert.That(fnotes, Is.Not.Null);
        }

        [Test]
        public void Empty_ShouldBeIgnored()
        {
            var elements = converter.Parse("<abbr></abbr>");
            Assert.That(elements, Is.Empty);
        }

        [TestCase("<abbr><a href='www.site.com'>Placeholder</a></abbr>")]
        [TestCase("<abbr>Placeholder</abbr>")]
        [TestCase("<blockquote>Placeholder</blockquote>")]
        public void WithNoDescription_ReturnsSimpleParagraph(string html)
        {
            // description nor title was defined - fallback to normal run
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Is.All.TypeOf<Paragraph>());
            Assert.That(mainPart.FootnotesPart, Is.Null);
            Assert.That(mainPart.EndnotesPart, Is.Null);
        }

        [TestCase("<abbr title='HyperText Markup Language'>HTML</abbr>", AcronymPosition.DocumentEnd, Description = "Read existing endnotes references")]
        [TestCase("<abbr title='HyperText Markup Language'>HTML</abbr>", AcronymPosition.PageEnd, Description = "Read existing footnotes references")]
        [TestCase("<blockquote cite='HyperText Markup Language'>HTML</blockquote>", AcronymPosition.DocumentEnd, Description = "Read existing endnotes references")]
        [TestCase("<blockquote cite='HyperText Markup Language'>HTML</blockquote>", AcronymPosition.PageEnd, Description = "Read existing footnotes references")]
        public void WithExistingEndnotes_ReturnsUniqueRefId(string html, AcronymPosition acronymPosition)
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithNotes.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;
            HtmlConverter converter = new(mainPart)
            {
                AcronymPosition = acronymPosition
            };

            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));

            FootnoteEndnoteReferenceType? noteRef;

            if (acronymPosition == AcronymPosition.PageEnd)
            {
                noteRef = elements[0].GetLastChild<Run>()?.GetFirstChild<FootnoteReference>();
                Assert.That(mainPart.FootnotesPart!.Footnotes.Elements<Footnote>().Select(fn => fn.Id?.Value), Is.Unique);
            }
            else
            {
                noteRef = elements[0].GetLastChild<Run>()?.GetFirstChild<EndnoteReference>();
                Assert.That(mainPart.EndnotesPart!.Endnotes.Elements<Endnote>().Select(fn => fn.Id?.Value), Is.Unique);
            }

            Assert.That(noteRef, Is.Not.Null);
            Assert.That(noteRef.Id?.HasValue, Is.EqualTo(true));

            FootnoteEndnoteType? note;
            if (acronymPosition == AcronymPosition.PageEnd)
            {
                note = mainPart.FootnotesPart!.Footnotes.Elements<Footnote>()
                    .FirstOrDefault(fn => fn.Id?.Value == noteRef.Id.Value);
            }
            else
            {
                note = mainPart.EndnotesPart!.Endnotes.Elements<Endnote>()
                    .FirstOrDefault(fn => fn.Id?.Value == noteRef.Id.Value);
            }
            Assert.That(note, Is.Not.Null);
            Assert.That(note.InnerText, Is.EqualTo(" " + "HyperText Markup Language"));
        }

        [Test]
        public void InsideParagraph_ReturnsMultipleRuns()
        {
            var elements = converter.Parse(@"<p>The 
                <abbr title='National Aeronautics and Space Administration'>NASA</abbr>
                is an independent agency of the U.S. federal government responsible for the civil space program, aeronautics research, and space research.</p>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.Multiple(() => {
                Assert.That(elements[0], Is.TypeOf(typeof(Paragraph)));
                Assert.That(elements[0].Elements<Run>().Count(), Is.EqualTo(6), "3 textual runs + 3 breaks");
                Assert.That(elements[0].Elements<Run>().Any(r => r.HasChild<FootnoteReference>()), Is.True);
            });
        }

        [TestCase("<abbr title='National Aeronautics and Space Administration'>NASA</abbr>")]
        [TestCase("<blockquote cite='https://en.wikipedia.org/wiki/NASA'>NASA</blockquote>")]
        public async Task ParseIntoHeader_ReturnsSimpleParagraph(string html)
        {
            await converter.ParseHeader(html);
            var header = mainPart.HeaderParts?.FirstOrDefault()?.Header;
            Assert.That(header, Is.Not.Null);
            Assert.That(header.ChildElements, Has.Count.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(header.ChildElements, Is.All.TypeOf<Paragraph>());
                Assert.That(header.FirstChild!.InnerText, Is.EqualTo("NASA"));
                Assert.That(mainPart.FootnotesPart, Is.Null);
                Assert.That(mainPart.EndnotesPart, Is.Null);
                AssertThatOpenXmlDocumentIsValid();
            });
        }

        [TestCase("<abbr title='National Aeronautics and Space Administration'>NASA</abbr>")]
        [TestCase("<blockquote cite='https://en.wikipedia.org/wiki/NASA'>NASA</blockquote>")]
        public async Task ParseIntoFooter_ShouldBeIgnored(string html)
        {
            await converter.ParseFooter(html);
            var footer = mainPart.FooterParts?.FirstOrDefault()?.Footer;
            Assert.That(footer, Is.Not.Null);
            Assert.That(footer.ChildElements, Has.Count.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(footer.ChildElements, Is.All.TypeOf<Paragraph>());
                Assert.That(footer.FirstChild!.InnerText, Is.EqualTo("NASA"));
                Assert.That(mainPart.FootnotesPart, Is.Null);
                Assert.That(mainPart.EndnotesPart, Is.Null);
                AssertThatOpenXmlDocumentIsValid();
            });
        }
    }
}