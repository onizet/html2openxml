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
font-family:Verdana;
font-variant:small-caps;
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
                Assert.That(runProperties.HasChild<Color>(), Is.True);
                Assert.That(runProperties.HasChild<SmallCaps>(), Is.True);
                Assert.That(runProperties.GetFirstChild<RunFonts>()?.Ascii?.Value, Is.EqualTo("Verdana"));
            });
        }

        [TestCase("<i style='font-style:normal'>Not italic</i>", false)]
        [TestCase("<span style='font-style:italic'><i style='font-style:normal'>Not italic</i></span>", false)]
        [TestCase("<span style='font-style:normal'><span style='font-style:italic'>Italic!</span></span>", true)]
        [TestCase("<div style='font-style:italic'><span style='font-style:normal'><span style='font-style:italic'>Italic!</span></span></div>", true)]
        [TestCase("<div style='font-style:italic'><div style='font-style:normal'>Not italic</div></div>", false)]
        [TestCase("<div id='outer' style='font-style:italic'><div id='inner'>Italic</div></div>", true)]
        public void ParseCascadeStyle (string html, bool expectItalic)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Is.Not.Empty);
            Assert.That(elements[0], Is.TypeOf<Paragraph>());
            Assert.That(elements[0].FirstChild, Is.TypeOf<Run>());
            var run  = elements[0].FirstChild as Run;
            if (expectItalic)
            {
                Assert.That(run.RunProperties, Is.Not.Null);
                Assert.That(run.RunProperties.Italic, Is.Not.Null);
                // normally, Val should be null
                if (run.RunProperties.Italic.Val is not null)
                    Assert.That(run.RunProperties.Italic.Val, Is.EqualTo(true));
            }
            else
            {
                // italic should not be applied as we specify font-style=normal
                if (run.RunProperties?.Italic is not null)
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
            Assert.That(elements[0].ChildElements, Has.Count.EqualTo(4));

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

        [TestCase(@"<span>Placeholder</span>")]
        [TestCase(@"<time datetime='2024-07-05'>5 July</time>")]
        public void ParseSimplePhrasing(string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));
            Run run = elements[0].GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);
            Assert.That(run.RunProperties, Is.Null);
        }

        [Test]
        public void ParseDefinitionList()
        {
            var elements = converter.Parse(@"
            <dl>
                <dt>Denim (semigloss finish)</dt>
                <dd>Ceiling</dd>
                <dt>Denim (eggshell finish)</dt>
                <dt>Evening Sky (eggshell finish)</dt>
                <dd>Layered on the walls</dd>
            </dl>");

            Assert.That(elements, Has.Count.EqualTo(5));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());

            var ddElements = elements.Where((e, idx) => idx == 1 || idx == 5);
            Assert.That(ddElements.All(p => p.GetFirstChild<ParagraphProperties>()?.HasChild<Indentation>() == true), Is.True,
                "All `dd` paragraph are converted with Indentation");
        }

        [Test]
        public void ParseAlternateL8ng()
        {
            var elements = converter.Parse(@"<div lang=""en"">
                <p>Pineapple</p>
                <p lang=""fr"">Ananas</p>
                <p lang=""ar"">أناناس</p>
                <p lang=""sindarin"">yávë</p>
            </div>");

            Assert.That(elements, Has.Count.EqualTo(4));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());

            int index = 0;
            foreach (var (rtl, lang) in new[] {
                (false, "en"), // inherited from parent container
                (false, "fr"),
                (true, "ar"),
                (false, "en") // unknown language - fallback on parent
            })
            {
                var p = (Paragraph) elements[index];
                Assert.That(p.ParagraphProperties?.BiDi?.Val?.Value, Is.EqualTo(rtl), $"{index}. expected RTL={rtl}");
                Assert.That(p.ParagraphProperties?.ParagraphMarkRunProperties?
                    .GetFirstChild<Languages>()?.Val?.Value, Is.EqualTo(lang), $"{index}. expected lang={lang}");
                Assert.That(p.GetFirstChild<Run>()?.GetFirstChild<RunProperties>()?
                    .Languages?.Val?.Value, Is.EqualTo(lang), $"{index}. expected lang={lang}");
                index++;
            }
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