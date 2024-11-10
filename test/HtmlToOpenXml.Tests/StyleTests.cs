using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests on custom styling.
    /// </summary>
    [TestFixture]
    public class StyleTests : HtmlConverterTestBase
    {
        [Test(Description = "Define on the fly a new Paragraph style and apply it")]
        public void UseVariantStyle_ReturnsAppliedStyle()
        {
            bool wasTriggered = false;
            converter.HtmlStyles.StyleMissing += delegate(object? sender, StyleEventArgs args) {
                if (args.Type != StyleValues.Paragraph)
                    return;
                wasTriggered = true;
                Assert.That(args.Name, Is.EqualTo("custom-style"));
                Assert.That(sender, Is.TypeOf<WordDocumentStyle>());
                ((WordDocumentStyle) sender).AddStyle(new Style() {
                    StyleId = "custom-style",
                    Type = args.Type,
                    BasedOn = new BasedOn { Val = "Normal" },
                    StyleRunProperties = new() {
                        Color = new() { Val = HtmlColorTranslator.FromHtml("red").ToHexString() }
                    }
                });
            };
            var elements = converter.Parse("<p class='custom-style'>Placeholder</p>");
            Assert.That(wasTriggered, Is.True);
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0], Is.TypeOf<Paragraph>());

            var paragraph = (Paragraph) elements[0];
            Assert.That(paragraph.ParagraphProperties, Is.Not.Null);
            Assert.That(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value, Is.EqualTo("custom-style"));
        }

        [Test(Description = "CustomStyle1 is defined in the provided document and must be discover")]
        public void UseVariantParagraphStyle_ReturnsAppliedStyle()
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithCustomStyle.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;
            HtmlConverter converter = new HtmlConverter(mainPart);

            var elements = converter.Parse("<div class='CustomStyle1'>Lorem</div><span>Ipsum</span>");
            Assert.That(elements, Is.Not.Empty);
            var paragraphProperties = elements[0].GetFirstChild<ParagraphProperties>();
            Assert.That(paragraphProperties, Is.Not.Null);
            Assert.That(paragraphProperties.ParagraphStyleId, Is.Not.Null);
            Assert.That(paragraphProperties.ParagraphStyleId?.Val?.Value, Is.EqualTo("CustomStyle1"));
        }

        [Test(Description = "TableNormal style define outside borders")]
        public void UseVariantTableStyle_ReturnsAppliedStyle()
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithCustomStyle.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;
            HtmlConverter converter = new HtmlConverter(mainPart);

            var elements = converter.Parse("<table class='TableNormal' border='2'><tr><td>Lorem Ipsum</td></tr></table>");
            Assert.That(elements, Is.Not.Empty);
            var tableProperties = elements[0].GetFirstChild<TableProperties>();
            Assert.That(tableProperties, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(tableProperties.TableStyle, Is.Not.Null);
                Assert.That(tableProperties.TableBorders, Is.Not.Null);
            });

            // the TableNormal doesn't define any borders while default shipped Table style does
            Assert.Multiple(() =>
            {
                Assert.That(tableProperties.TableStyle?.Val?.Value, Is.EqualTo("TableNormal"));
                Assert.That(tableProperties.TableBorders.LeftBorder?.Val?.Value, Is.EqualTo(BorderValues.None));
                Assert.That(tableProperties.TableBorders.TopBorder?.Val?.Value, Is.EqualTo(BorderValues.None));
                Assert.That(tableProperties.TableBorders.RightBorder?.Val?.Value, Is.EqualTo(BorderValues.None));
                Assert.That(tableProperties.TableBorders.BottomBorder?.Val?.Value, Is.EqualTo(BorderValues.None));
                Assert.That(tableProperties.TableBorders.InsideHorizontalBorder?.Val?.Value, Is.EqualTo(BorderValues.Single));
                Assert.That(tableProperties.TableBorders.InsideVerticalBorder?.Val?.Value, Is.EqualTo(BorderValues.Single));
            });
        }

        [Test]
        public void SetDefaultStyle_Paragraph_ReturnsAppliedStyle()
        {
            converter.HtmlStyles.DefaultStyles.IntenseQuoteStyle = "CustomIntenseQuoteStyle";
            converter.HtmlStyles.AddStyle(new Style {
                StyleId = "CustomIntenseQuoteStyle",
                Type = StyleValues.Paragraph,
                StyleParagraphProperties = new() {
                    ParagraphBorders = new() {
                        LeftBorder = new() { Val = BorderValues.Single, Color = HtmlColor.FromArgb(255, 0, 0).ToHexString() }
                    }
                }
            });

            bool wasTriggered = false;
            converter.HtmlStyles.StyleMissing += delegate(object? sender, StyleEventArgs args) {
                wasTriggered = true;
                Assert.That(args.Type, Is.EqualTo(StyleValues.Paragraph));
                Assert.That(args.Name, Is.EqualTo("CustomIntenseQuoteStyle"));
            };
            var elements = converter.Parse(@"<blockquote cite=""http://www.worldwildlife.org/who/index.html"">
For 50 years, <b>WWF</b> has been protecting the future of nature. The world's leading conservation organization, WWF works in 100 countries and is supported by 1.2 million members in the United States and close to 5 million globally.
</blockquote> ");
            Assert.That(wasTriggered, Is.False);
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0], Is.TypeOf<Paragraph>());

            var paragraph = (Paragraph) elements[0];
            Assert.That(paragraph.ParagraphProperties, Is.Not.Null);
            Assert.That(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value, Is.EqualTo("CustomIntenseQuoteStyle"));
        }

        [Test(Description = "Appending style into StyleDefinionsPart requires a call to RefreshStyles")]
        public void ManualAddStyle_ThenRefreshStyles_ShouldSucceed()
        {
            var stylePart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles ??= new();
            stylePart.Styles.AddChild(new Style {
                Type = StyleValues.Paragraph,
                StyleId = "CustomIntenseQuoteStyle",
                StyleName = new() { Val = "CustomIntenseQuoteStyle" },
                StyleHidden = new() { Val = OnOffOnlyValues.On },
                StyleParagraphProperties = new() {
                    ParagraphBorders = new() {
                        LeftBorder = new() { Val = BorderValues.Single, Color = HtmlColor.FromArgb(255, 0, 0).ToHexString() }
                    }
                }
            });
            converter.RefreshStyles();

            bool wasTriggered = false;
            converter.HtmlStyles.StyleMissing += delegate(object? sender, StyleEventArgs args) {
                if (args.Name == "CustomIntenseQuoteStyle" && args.Type == StyleValues.Paragraph) {
                    wasTriggered = true;
                }
            };
            var elements = converter.Parse(@"<blockquote class=""CustomIntenseQuoteStyle"" cite=""http://www.worldwildlife.org/who/index.html"">
For 50 years, <b>WWF</b> has been protecting the future of nature. The world's leading conservation organization, WWF works in 100 countries and is supported by 1.2 million members in the United States and close to 5 million globally.
</blockquote> ");
            Assert.That(wasTriggered, Is.False);
            var paragraph = (Paragraph) elements[0];
            Assert.That(paragraph.ParagraphProperties, Is.Not.Null);
            Assert.That(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value, Is.EqualTo("CustomIntenseQuoteStyle"));
        }

        [Test(Description = "Parser should consider the last occurence of a style")]
        public void DuplicateStyle_ReturnsLatter()
        {
            var styleAttributes = HtmlAttributeCollection.ParseStyle("color:red;color:blue");
            Assert.That(styleAttributes["color"], Is.EqualTo("blue"));
        }

        [Test(Description = "Encoded ':' and ';' characters are valid")]
        public void EncodedStyle_ShouldSucceed()
        {
            var styleAttributes = HtmlAttributeCollection.ParseStyle("text-decoration&#58;underline&#59;color:red");
            Assert.That(styleAttributes["text-decoration"], Is.EqualTo("underline"));
            Assert.That(styleAttributes["color"], Is.EqualTo("red"));
        }

        [Test(Description = "Key style with no value should be ignored")]
        public void EmptyStyle_ShouldBeIgnoredd()
        {
            var styleAttributes = HtmlAttributeCollection.ParseStyle("text-decoration;color:red");
            Assert.That(styleAttributes["text-decoration"], Is.Null);
            Assert.That(styleAttributes["color"], Is.EqualTo("red"));
        }
    }
}
