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

        [TestCase(@"<sub>Subscript</sub>", "subscript")]
        [TestCase(@"<sup>Superscript</sup>", "superscript")]
        public void ParseSubSup (string html, string tagName)
        {
            var val = new VerticalPositionValues(tagName);
            var textAlign = ParsePhrasing<VerticalTextAlignment>(html);
            Assert.That(textAlign.Val.HasValue, Is.True);
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
            Assert.That(elements, Has.Count.EqualTo(1));

            Run run = elements[0].GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);

            RunProperties runProperties = run.GetFirstChild<RunProperties>();
            Assert.That(runProperties, Is.Not.Null);
            Assert.Multiple(() => {
                Assert.That(runProperties.HasChild<Bold>(), Is.True);
                Assert.That(runProperties.HasChild<Italic>(), Is.True);
                Assert.That(runProperties.HasChild<FontSize>(), Is.True);
                Assert.That(runProperties.HasChild<Underline>(), Is.True);
                Assert.That(runProperties.HasChild<Color>(),Is.True);
            });
        }

        [TestCase(@"<i style='font-style:normal'>Not italic</i>", false)]
        [TestCase(@"<span style='font-style:italic'><i style='font-style:normal'>Not italic</i></span>", false)]
        [TestCase(@"<span style='font-style:normal'><span style='font-style:italic'>Italic!</span></span>", true)]
        public void ParseDisruptiveStyle (string html, bool expectItalic)
        {
            // italic should not be applied as we specify font-style=normal
            var elements = converter.Parse(html);
            Assert.That(elements, Is.Not.Empty);
            Assert.That(elements[0], Is.TypeOf<Paragraph>());
            Assert.That(elements[0].FirstChild, Is.TypeOf<Run>());
            var run  = elements[0].FirstChild as Run;
            Assert.That(run.RunProperties, Is.Not.Null);
            if (expectItalic)
            {
                Assert.That(run.RunProperties.Italic, Is.Not.Null);
                // normally, Val should be null
                if (run.RunProperties.Italic.Val is not null)
                    Assert.That(run.RunProperties.Italic.Val, Is.EqualTo(true));
            }
            else
            {
                if (run.RunProperties.Italic is not null)
                    Assert.That(run.RunProperties.Italic.Val, Is.EqualTo(false));
            }
        }

        [TestCase(@"<q>Build a future where people live in harmony with nature.</q>", true)]
        [TestCase(@"<cite>Build a future where people live in harmony with nature.</cite>", false)]
        public void ParseQuote(string html, bool hasQuote)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));

            Run run = elements[0].GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);
            if (hasQuote)
            {
                Assert.That(run.InnerText, Is.EqualTo(" " + converter.HtmlStyles.QuoteCharacters.Prefix));

                Run lastRun = elements[0].GetLastChild<Run>();
                Assert.That(run, Is.Not.Null);
                Assert.That(lastRun.InnerText, Is.EqualTo(converter.HtmlStyles.QuoteCharacters.Suffix));

                // focus the content run
                run = (Run) run.NextSibling();
            }

            RunProperties runProperties = run.GetFirstChild<RunProperties>();
            Assert.That(runProperties, Is.Not.Null);

            var runStyle = runProperties.GetFirstChild<RunStyle>();
            Assert.That(runStyle, Is.Not.Null);
            Assert.That(runStyle.Val.Value, Is.EqualTo("QuoteChar"));
        }

        [Test]
        public void ParseBreak()
        {
            var elements = converter.Parse(@"Lorem<br/>Ipsum");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0].ChildElements, Has.Count.EqualTo(3));

            Assert.Multiple(() =>
            {
                Assert.That(elements[0].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(((Run)elements[0].ChildElements[1]).GetFirstChild<Break>(), Is.Not.Null);
            });
        }

        [Test]
        public void ParseFigCaption()
        {
            var elements = converter.Parse(@"<figcaption>Fig.1 - Trulli, Puglia, Italy.</figcaption>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0], Is.TypeOf<Paragraph>());

            Assert.Multiple(() =>
            {
                Assert.That(elements[0].ChildElements, Has.Count.EqualTo(3));
                Assert.That(elements[0].HasChild<Run>(), Is.True);
                Assert.That(elements[0].HasChild<SimpleField>(), Is.True);
            });
        }

        [Test]
        public void ParseFont ()
        {
            var elements = converter.Parse(@"<font size=""small"" face=""Verdana"">Placeholder</font>");
            Assert.That(elements, Has.Count.EqualTo(1));
            var run = elements[0].GetFirstChild<Run>();
            Assert.Multiple(() => {
                Assert.That(run.RunProperties.FontSize, Is.Not.Null);
                Assert.That(run.RunProperties.RunFonts?.Ascii?.Value, Is.EqualTo("Verdana"));
            });
        }

        private T ParsePhrasing<T> (string html) where T : OpenXmlElement
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));

            Run run = elements[0].GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);

            RunProperties runProperties = run.GetFirstChild<RunProperties>();
            Assert.That(runProperties, Is.Not.Null);

            var tag = runProperties.GetFirstChild<T>();
            Assert.That(tag, Is.Not.Null);
            return tag;
        }
    }
}