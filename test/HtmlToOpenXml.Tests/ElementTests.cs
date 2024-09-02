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
        public void PhrasingTag_ReturnsRunWithDefaultStyle<T> (string html) where T : OpenXmlElement
        {
            Assert.That(ParsePhrasing<T>(html), Is.TypeOf<T>());
        }

        [TestCase(@"<sub>Subscript</sub>", ExpectedResult = "subscript")]
        [TestCase(@"<sup>Superscript</sup>", ExpectedResult = "superscript")]
        public string? SubSup_ReturnsRunWithVerticalAlignment (string html)
        {
            //var val = new VerticalPositionValues(tagName);
            var textAlign = ParsePhrasing<VerticalTextAlignment>(html);
            Assert.That(textAlign.Val?.HasValue, Is.True);
            return textAlign.Val.InnerText;
        }

        [Test]
        public void MultipleStyle_ShouldBeAllApplied ()
        {
            var elements = converter.Parse(@"<b style=""
font-style:italic;
font-size:12px;
font-family:Verdana;
font-variant:small-caps;
color:white;
text-decoration:wavy line-through double;
background:red;
"">bold with italic style</b>");
            Assert.That(elements, Has.Count.EqualTo(1));

            var run = elements[0].GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);

            var runProperties = run.GetFirstChild<RunProperties>();
            Assert.That(runProperties, Is.Not.Null);
            Assert.Multiple(() => {
                Assert.That(runProperties.HasChild<Bold>(), Is.True);
                Assert.That(runProperties.HasChild<Italic>(), Is.True);
                Assert.That(runProperties.HasChild<FontSize>(), Is.True);
                Assert.That(runProperties.HasChild<Underline>(), Is.True);
                Assert.That(runProperties.HasChild<DoubleStrike>(), Is.True);
                Assert.That(runProperties.HasChild<SmallCaps>(), Is.True);
                Assert.That(runProperties.Color?.Val?.Value, Is.EqualTo("FFFFFF"));
                Assert.That(runProperties.RunFonts?.Ascii?.Value, Is.EqualTo("Verdana"));
                Assert.That(runProperties.Underline?.Val?.Value, Is.EqualTo(UnderlineValues.Wave));
                Assert.That(runProperties.Shading?.Fill?.Value, Is.EqualTo("FF0000"));
            });
        }

        [TestCase("<span style='font-style:normal'><span style='font-style:italic'>Italic!</span></span>")]
        [TestCase("<div style='font-style:italic'><span style='font-style:normal'><span style='font-style:italic'>Italic!</span></span></div>")]
        [TestCase("<div id='outer' style='font-style:italic'><div id='inner'>Italic</div></div>")]
        public void NestedTagWithStyle_ShouldCascadeParentStyle (string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Is.Not.Empty);
            Assert.That(elements[0], Is.TypeOf<Paragraph>());
            Assert.That(elements[0].FirstChild, Is.TypeOf<Run>());
            var run = elements[0].FirstChild as Run;

            Assert.That(run?.RunProperties, Is.Not.Null);
            Assert.That(run.RunProperties.Italic, Is.Not.Null);
            // normally, Val should be null
            if (run.RunProperties.Italic.Val is not null)
                Assert.That(run.RunProperties.Italic.Val, Is.EqualTo(true));
        }

        [TestCase("<i style='font-style:normal'>Not italic</i>")]
        [TestCase("<span style='font-style:italic'><i style='font-style:normal'>Not italic</i></span>")]
        [TestCase("<div style='font-style:italic'><div style='font-style:normal'>Not italic</div></div>")]
        public void NestedTagWithStyle_ShouldOverrideParentStyle (string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Is.Not.Empty);
            Assert.That(elements[0], Is.TypeOf<Paragraph>());
            Assert.That(elements[0].FirstChild, Is.TypeOf<Run>());
            var run = elements[0].FirstChild as Run;

            // italic should not be applied as we specify font-style=normal
            if (run?.RunProperties?.Italic is not null)
                Assert.That(run.RunProperties.Italic.Val, Is.EqualTo(false));
        }

        [TestCase(@"<q>Build a future where people live in harmony with nature.</q>", true)]
        [TestCase(@"<quote>Build a future where people live in harmony with nature.</quote>", true)]
        [TestCase(@"<cite>Build a future where people live in harmony with nature.</cite>", false)]
        public void Quote_ReturnsRunWithStyleAndFormat(string html, bool hasQuote)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));

            var run = elements[0].GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);
            if (hasQuote)
            {
                Assert.That(run.InnerText, Is.EqualTo(" " + converter.HtmlStyles.QuoteCharacters.Prefix));

                var lastRun = elements[0].GetLastChild<Run>();
                Assert.That(run, Is.Not.Null);
                Assert.That(lastRun?.InnerText, Is.EqualTo(converter.HtmlStyles.QuoteCharacters.Suffix));

                // focus the content run
                run = (Run?) run.NextSibling();
            }

            var runProperties = run.GetFirstChild<RunProperties>();
            Assert.That(runProperties, Is.Not.Null);

            var runStyle = runProperties.GetFirstChild<RunStyle>();
            Assert.That(runStyle, Is.Not.Null);
            Assert.That(runStyle.Val?.Value, Is.EqualTo("QuoteChar"));
        }

        [Test]
        public void TextWithBreak_ReturnsRunWithBreak()
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
        public void FigCaption_ReturnsRunWithSimpleField()
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
        public void FontFamily_ReturnsRunWithFont ()
        {
            var elements = converter.Parse(@"<font size=""small"" face=""Verdana"">Placeholder</font>");
            Assert.That(elements, Has.Count.EqualTo(1));
            var run = elements[0].GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);
            Assert.Multiple(() => {
                Assert.That(run.RunProperties?.FontSize, Is.Not.Null);
                Assert.That(run.RunProperties?.RunFonts?.Ascii?.Value, Is.EqualTo("Verdana"));
            });
        }

        [TestCase(@"<span>Placeholder</span>")]
        [TestCase(@"<time datetime='2024-07-05'>5 July</time>")]
        public void PhrasingTag_ReturnsRunWithText(string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));
            var run = elements[0].GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);
            Assert.That(run.RunProperties, Is.Null);
        }

        [Test]
        public void DefinitionList_ReturnsIndentedParagraphs()
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

        [TestCase("<p lang='fr'>Ananas</p>", false, "fr")]
        [TestCase("<p lang='ar'>أناناس</p>", true, "ar")]
        public void AlternateL8ng_ReturnsSpecificBidi(string paragraphHtml, bool expectRtl, string expectLang)
        {
            var elements = converter.Parse($@"<div lang=""en"">{paragraphHtml}</div>");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());

            var p = (Paragraph) elements[0];
            Assert.Multiple(() =>
            {
                Assert.That(p.ParagraphProperties?.BiDi?.Val?.Value, Is.EqualTo(expectRtl), $"Expected RTL={expectRtl}");
                Assert.That(p.ParagraphProperties?.ParagraphMarkRunProperties?
                    .GetFirstChild<Languages>()?.Val?.Value, Is.EqualTo(expectLang), $"expected lang={expectLang}");
                Assert.That(p.GetFirstChild<Run>()?.GetFirstChild<RunProperties>()?
                    .Languages?.Val?.Value, Is.EqualTo(expectLang), $"expected lang={expectLang}");
            });
        }

        [TestCase("<p>Pineapple</p>", Description = "Inherited from parent container")]
        [TestCase("<p lang='sindarin'>yávë</p>", Description = "Unknown language -> fallback on parent")]
        public void Failed_AlternateL8ng_ReturnsInheritedBidi(string paragraphHtml)
        {
            var elements = converter.Parse($@"<div lang=""en"">{paragraphHtml}</div>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());

            var p = (Paragraph) elements[0];
            Assert.Multiple(() =>
            {
                Assert.That(p.ParagraphProperties?.BiDi?.Val?.Value, Is.EqualTo(false));
                Assert.That(p.ParagraphProperties?.ParagraphMarkRunProperties?
                    .GetFirstChild<Languages>()?.Val?.Value, Is.EqualTo("en"));
                Assert.That(p.GetFirstChild<Run>()?.GetFirstChild<RunProperties>()?
                    .Languages?.Val?.Value, Is.EqualTo("en"));
            });
        }

        private T ParsePhrasing<T> (string html) where T : OpenXmlElement
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));

            var run = elements[0].GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);

            var runProperties = run.GetFirstChild<RunProperties>();
            Assert.That(runProperties, Is.Not.Null);

            var tag = runProperties.GetFirstChild<T>();
            Assert.That(tag, Is.Not.Null);
            return tag;
        }
    }
}