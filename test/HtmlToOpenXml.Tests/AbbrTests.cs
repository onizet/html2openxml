using System.Linq;
using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests acronym, abbreviation and blockquotes.
    /// </summary>
    [TestFixture]
    public class AbbrTests : HtmlConverterTestBase
    {
        [TestCase(@"<abbr title='National Aeronautics and Space Administration'>NASA</abbr>")]
        [TestCase(@"<acronym title='National Aeronautics and Space Administration'>NASA</acronym>")]
        [TestCase(@"<acronym title='www.nasa.gov'>NASA</acronym>")]
        public void ParseAbbr(string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements.Count, Is.EqualTo(1));
            Assert.Multiple(() => {
                Assert.That(elements[0], Is.TypeOf(typeof(Paragraph)));
                Assert.That(elements[0].LastChild, Is.TypeOf(typeof(Run)));
                Assert.That(elements[0].InnerText, Is.EqualTo("NASA"));
            });

            var noteRef = elements[0].LastChild.GetFirstChild<FootnoteReference>();
            Assert.IsNotNull(noteRef);
            Assert.That(noteRef.Id.HasValue, Is.EqualTo(true));

            Assert.IsNotNull(mainPart.FootnotesPart);
            Assert.That(mainPart.FootnotesPart.HyperlinkRelationships.Count(), Is.EqualTo(0));

            var fnotes = mainPart.FootnotesPart.Footnotes.Elements<Footnote>().FirstOrDefault(f => f.Id.Value == noteRef.Id.Value);
            Assert.IsNotNull(fnotes);
        }

        [TestCase(@"<abbr title='https://en.wikipedia.org/wiki/N A S A '>NASA</abbr>", "https://en.wikipedia.org/wiki/N%20A%20S%20A")]
        [TestCase(@"<abbr title='file://C:\temp\NASA.html'>NASA</abbr>", @"file:///C:/temp/NASA.html")]
        [TestCase(@"<abbr title='\\server01\share\NASA.html'>NASA</abbr>", "file://server01/share/NASA.html")]
        [TestCase(@"<abbr title='ftp://server01/share/NASA.html'>NASA</abbr>", "ftp://server01/share/NASA.html")]
        [TestCase(@"<blockquote cite='https://en.wikipedia.org/wiki/NASA'>NASA</blockquote>", "https://en.wikipedia.org/wiki/NASA")]
        public void ParseWithLinks(string html, string expectedUri)
        {
            var elements = converter.Parse(html);
            Assert.That(elements.Count, Is.EqualTo(1));
            Assert.Multiple(() => {
                Assert.That(elements[0], Is.TypeOf(typeof(Paragraph)));
                Assert.That(elements[0].LastChild, Is.TypeOf(typeof(Run)));
                Assert.That(elements[0].InnerText, Is.EqualTo("NASA"));
            });

            var noteRef = elements[0].LastChild.GetFirstChild<FootnoteReference>();
            Assert.IsNotNull(noteRef);
            Assert.That(noteRef.Id.HasValue, Is.EqualTo(true));

            Assert.IsNotNull(mainPart.FootnotesPart);
            var fnotes = mainPart.FootnotesPart.Footnotes.Elements<Footnote>().FirstOrDefault(f => f.Id.Value == noteRef.Id.Value);
            Assert.IsNotNull(fnotes);

            var link = fnotes.FirstChild.GetFirstChild<Hyperlink>();
            Assert.IsNotNull(link);

            var extLink = mainPart.FootnotesPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == link.Id);
            Assert.IsNotNull(extLink);
            Assert.That(extLink.IsExternal, Is.EqualTo(true));
            Assert.That(extLink.Uri.AbsoluteUri, Is.EqualTo(expectedUri));
        }

        [Test]
        public void ParseDocumentEnd()
        {
            converter.AcronymPosition = AcronymPosition.DocumentEnd;
            var elements = converter.Parse(@"<acronym title='www.nasa.gov'>NASA</acronym>");

            var noteRef = elements[0].LastChild.GetFirstChild<EndnoteReference>();
            Assert.IsNotNull(noteRef);
            Assert.That(noteRef.Id.HasValue, Is.EqualTo(true));

            Assert.IsNotNull(mainPart.EndnotesPart);
            var fnotes = mainPart.EndnotesPart.Endnotes.Elements<Endnote>().FirstOrDefault(f => f.Id.Value == noteRef.Id.Value);
            Assert.IsNotNull(fnotes);
        }
    }
}