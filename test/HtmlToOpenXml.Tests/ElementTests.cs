using NUnit.Framework;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests Bold, Italic, Underline, Strikethrough.
    /// </summary>
    [TestFixture]
    public class ElementTests : HtmlConverterTestBase
    {
        [GenericTestCase(typeof(Bold), @"<b>Bold</b>")]
        [GenericTestCase(typeof(Bold), @"<strong>Strong</strong>")]
        [GenericTestCase(typeof(Italic), @"<i>Italic</i>")]
        [GenericTestCase(typeof(Italic), @"<em>Italic</em>")]
        [GenericTestCase(typeof(Strike), @"<s>Strike</s>")]
        [GenericTestCase(typeof(Strike), @"<strike>Strike</strike>")]
        [GenericTestCase(typeof(Strike), @"<del>Del</del>")]
        [GenericTestCase(typeof(Underline), @"<u>Underline</u>")]
        [GenericTestCase(typeof(Underline), @"<ins>Inserted</ins>")]
        public void ParseHtmlElements<T> (string html) where T : OpenXmlElement
        {
            ParsePhrasing<T>(html);
        }

        [TestCase(@"<sub>Subscript</sub>", VerticalPositionValues.Subscript)]
        [TestCase(@"<sup>Superscript</sup>", VerticalPositionValues.Superscript)]
        public void ParseSubSup (string html, VerticalPositionValues val)
        {
            var textAlign = ParsePhrasing<VerticalTextAlignment>(html);
            Assert.That(textAlign.Val.HasValue, Is.EqualTo(true));
            Assert.That(textAlign.Val.Value, Is.EqualTo(val));
        }

        [Test]
        public void ParseStyle ()
        {
            var elements = converter.Parse(@"<b style=""
font-style:italic;
font-size:12px;
color:red;
text-decoration:underline;
"">bold with italic style</b>");
            Assert.That(elements.Count, Is.EqualTo(1));

            Run run = elements[0].GetFirstChild<Run>();
            Assert.IsNotNull(run);

            RunProperties runProperties = run.GetFirstChild<RunProperties>();
            Assert.IsNotNull(runProperties);
            Assert.Multiple(() => {
                Assert.IsTrue(runProperties.HasChild<Bold>());
                Assert.IsTrue(runProperties.HasChild<Italic>());
                Assert.IsTrue(runProperties.HasChild<FontSize>());
                Assert.IsTrue(runProperties.HasChild<Underline>());
                Assert.IsTrue(runProperties.HasChild<Color>());
            });
        }

        /*[Test]
        public void ParseDisruptiveStyle ()
        {
            //TODO:
            // italic should not be applied as we specify font-style=normal
            var elements = converter.Parse("<i style='font-style:normal'>Not italics</i>");
            Assert.Multiple(() => {
                var runProperties = elements[0].FirstChild.GetFirstChild<RunProperties>();
                Assert.IsNotNull(runProperties);
                Assert.IsTrue(!runProperties.HasChild<Italic>());
            });

            elements = converter.Parse("<span style='font-style:italic'><i style='font-style:normal'>Not italics</i></span>");
        }*/

        [TestCase(@"<q>Build a future where people live in harmony with nature.</q>", true)]
        [TestCase(@"<cite>Build a future where people live in harmony with nature.</cite>", false)]
        public void ParseQuote(string html, bool hasQuote)
        {
            var elements = converter.Parse(html);
            Assert.That(elements.Count, Is.EqualTo(1));

            Run run = elements[0].GetFirstChild<Run>();
            Assert.IsNotNull(run);
            if (hasQuote)
            {
                Assert.That(run.InnerText, Is.EqualTo(" " + converter.HtmlStyles.QuoteCharacters.Prefix));

                Run lastRun = elements[0].GetLastChild<Run>();
                Assert.IsNotNull(run);
                Assert.That(lastRun.InnerText, Is.EqualTo(converter.HtmlStyles.QuoteCharacters.Suffix));

                // focus the content run
                run = (Run) run.NextSibling();
            }

            RunProperties runProperties = run.GetFirstChild<RunProperties>();
            Assert.IsNotNull(runProperties);

            var runStyle = runProperties.GetFirstChild<RunStyle>();
            Assert.IsNotNull(runStyle);
            Assert.That(runStyle.Val.Value, Is.EqualTo("QuoteChar"));
        }

        [Test]
        public void ParseBreak()
        {
            var elements = converter.Parse(@"Lorem<br/>Ipsum");
            Assert.That(elements.Count, Is.EqualTo(1));
            Assert.That(elements[0].ChildElements.Count, Is.EqualTo(3));

            Assert.That(elements[0].ChildElements[0], Is.InstanceOf(typeof(Run)));
            Assert.That(elements[0].ChildElements[1], Is.InstanceOf(typeof(Run)));
            Assert.That(elements[0].ChildElements[2], Is.InstanceOf(typeof(Run)));
            Assert.IsNotNull(((Run)elements[0].ChildElements[1]).GetFirstChild<Break>());
        }

        private T ParsePhrasing<T> (string html) where T : OpenXmlElement
        {
            var elements = converter.Parse(html);
            Assert.That(elements.Count, Is.EqualTo(1));

            Run run = elements[0].GetFirstChild<Run>();
            Assert.IsNotNull(run);

            RunProperties runProperties = run.GetFirstChild<RunProperties>();
            Assert.IsNotNull(runProperties);

            var tag = runProperties.GetFirstChild<T>();
            Assert.IsNotNull(tag);
            return tag;
        }
    }
}