using System;
using System.Linq;
using NUnit.Framework;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests parser with various complex input Html.
    /// </summary>
    [TestFixture]
    public class ParserTests : HtmlConverterTestBase
    {
        [TestCase("<!--<p>some text</p>-->")]
        [TestCase("<script>$.appendTo('<p>some text</p>', document);</script>")]
        public void ParseIgnore(string html)
        {
            // the inner html shouldn't be interpreted
            var elements = converter.Parse(html);
            Assert.That(elements.Count, Is.EqualTo(0));
        }

        [Test]
        public void ParseUnclosedTag()
        {
            var elements = converter.Parse("<p>some text in <i>italics <b>,bold and italics</p>");
            Assert.Multiple(() => {
                Assert.That(elements.Count, Is.EqualTo(1));
                Assert.That(elements[0].ChildElements.Count, Is.EqualTo(3));

                var runProperties = elements[0].ChildElements[0].GetFirstChild<RunProperties>();
                Assert.IsNull(runProperties);

                runProperties = elements[0].ChildElements[1].GetFirstChild<RunProperties>();
                Assert.IsNotNull(runProperties);
                Assert.That(runProperties.HasChild<Italic>(), Is.EqualTo(true));
                Assert.That(runProperties.HasChild<Bold>(), Is.EqualTo(false));

                runProperties = elements[0].ChildElements[2].GetFirstChild<RunProperties>();
                Assert.IsNotNull(runProperties);
                Assert.That(runProperties.HasChild<Italic>(), Is.EqualTo(true));
                Assert.That(runProperties.HasChild<Bold>(), Is.EqualTo(true));
            });

            elements = converter.Parse("<p>First paragraph in semi-<i>italics <p>Second paragraph still italic <b>but also in bold</b></p>");
            Assert.Multiple(() => {
                Assert.That(elements.Count, Is.EqualTo(2));
                Assert.That(elements[0].ChildElements.Count, Is.EqualTo(2));
                Assert.That(elements[1].ChildElements.Count, Is.EqualTo(2));

                var runProperties = elements[0].ChildElements[0].GetFirstChild<RunProperties>();
                Assert.IsNull(runProperties);

                runProperties = elements[0].ChildElements[1].GetFirstChild<RunProperties>();
                Assert.IsNotNull(runProperties);
                Assert.That(runProperties.HasChild<Italic>(), Is.EqualTo(true));

                runProperties = elements[1].FirstChild.GetFirstChild<RunProperties>();
                Assert.IsNotNull(runProperties);
                Assert.That(runProperties.HasChild<Italic>(), Is.EqualTo(true));
                Assert.That(runProperties.HasChild<Bold>(), Is.EqualTo(false));

                runProperties = elements[1].ChildElements[1].GetFirstChild<RunProperties>();
                Assert.IsNotNull(runProperties);
                Assert.That(runProperties.HasChild<Italic>(), Is.EqualTo(true));
                Assert.That(runProperties.HasChild<Bold>(), Is.EqualTo(true));
            });

            // this should generate a new paragraph with its own style
            elements = converter.Parse("<p>First paragraph in <i>italics </i><p>Second paragraph not in italic</p>");
            Assert.Multiple(() => {
                Assert.That(elements.Count, Is.EqualTo(2));
                Assert.That(elements[0].ChildElements.Count, Is.EqualTo(2));
                Assert.That(elements[1].ChildElements.Count, Is.EqualTo(1));
                Assert.That(elements[1].FirstChild, Is.TypeOf(typeof(Run)));

                var runProperties = elements[1].FirstChild.GetFirstChild<RunProperties>();
                Assert.IsNull(runProperties);
            });
        }

        [TestCase("<p>Some\ntext</p>", ExpectedResult = 1)]
        [TestCase("<p>Some <b>bold\n</b>text</p>", ExpectedResult = 3)]
        [TestCase("\t<p>Some <b>bold\n</b>text</p>", ExpectedResult = 3)]
        [TestCase("  <p>Some text</p> ", ExpectedResult = 1)]
        public int ParseNewline (string html)
        {
            var elements = converter.Parse(html);
            return elements[0].ChildElements.Count;
        }

        [Test]
        public void ParseDisorderedTable ()
        {
            // table parts should be reordered
            var elements = converter.Parse(@"
<table>
<tbody>
    <tr><td>Body</td></tr>
</tbody>
<thead>
    <tr><td>Header</td></tr>
</thead>
<tfoot>
    <tr><td>Footer</td></tr>
</tfoot>
</table>");

            Assert.Multiple(() => {
                Assert.That(elements.Count, Is.EqualTo(1));
                Assert.That(elements[0], Is.TypeOf(typeof(Table)));

                var rows = elements[0].Elements<TableRow>();
                Assert.That(rows.Count(), Is.EqualTo(3));
                Assert.That(rows.ElementAt(0).InnerText, Is.EqualTo("Header"));
                Assert.That(rows.ElementAt(1).InnerText, Is.EqualTo("Body"));
                Assert.That(rows.ElementAt(2).InnerText, Is.EqualTo("Footer"));
            });
        }

        [Test]
        public void ParseNotTag ()
        {
            var elements = converter.Parse(" < b >bold</b>");
            Assert.Multiple(() => {
                Assert.That(elements.Count, Is.EqualTo(1));
                Assert.That(elements[0].ChildElements.Count, Is.EqualTo(1));
                Assert.IsNull(elements[0].FirstChild.GetFirstChild<RunProperties>());
            });

            elements = converter.Parse(" <3");
            Assert.Multiple(() => {
                Assert.That(elements.Count, Is.EqualTo(1));
                Assert.That(elements[0].ChildElements.Count, Is.EqualTo(1));
                Assert.IsNull(elements[0].FirstChild.GetFirstChild<RunProperties>());
            });
        }

        [Test]
        public void ParseNewlineFlow ()
        {
            // the new line should generate a space between "bold" and "text"
            var elements = converter.Parse(" <span>This is a <b>bold\n</b>text</span>");
        }
    }
}