using NUnit.Framework;

namespace HtmlToOpenXml.Tests.Primitives
{
    /// <summary>
    /// Tests Html font style attribute.
    /// </summary>
    [TestFixture]
    public class FontTests
    {
        [TestCase("1.2em Verdana", ExpectedResult = true)]
        [TestCase("Verdana 1.2em", ExpectedResult = false)]
        [TestCase("italic Verdana", ExpectedResult = false)]
        public bool WithMinimal_ReturnsValid (string html)
        {
            var font = HtmlFont.Parse(html);
            Assert.Multiple(() => {
                Assert.That(font.Style, Is.Null);
                Assert.That(font.Weight, Is.Null);
            });
            return font.Size.IsValid;
        }

        [TestCase("italic BOLD 1.2em Verdana")]
        [TestCase("Verdana  1.2em  bold  italic ")]
        public void WithDisordered_ShouldSucceed (string html)
        {
            var font = HtmlFont.Parse(html);
            Assert.Multiple(() => {
                Assert.That(font.Style, Is.EqualTo(FontStyle.Italic));
                Assert.That(font.Weight, Is.EqualTo(FontWeight.Bold));
                Assert.That(font.Family, Is.EqualTo("Verdana"));
                Assert.That(font.Size.Metric, Is.EqualTo(UnitMetric.EM));
                Assert.That(font.Size.Value, Is.EqualTo(1.2));
            });
        }

        [Test(Description = "Multiple font families must keep the first one")]
        public void WithMultipleFamily_ShouldSucceed ()
        {
            var font = HtmlFont.Parse("Verdana, Arial bolder 1.2em");
            Assert.Multiple(() => {
                Assert.That(font.Style, Is.Null);
                Assert.That(font.Weight, Is.EqualTo(FontWeight.Bolder));
                Assert.That(font.Family, Is.EqualTo("Verdana"));
                Assert.That(font.Size.Metric, Is.EqualTo(UnitMetric.EM));
                Assert.That(font.Size.Value, Is.EqualTo(1.2));
            });
        }

        [Test(Description = "Font families with quotes must unescape the first one")]
        public void WithQuotedFamily_ShouldSucceed ()
        {
            var font = HtmlFont.Parse("'Times New Roman', Times, Verdana, Arial bolder 1.2em");
            Assert.Multiple(() => {
                Assert.That(font.Style, Is.Null);
                Assert.That(font.Weight, Is.EqualTo(FontWeight.Bolder));
                Assert.That(font.Family, Is.EqualTo("Times New Roman"));
                Assert.That(font.Size.Metric, Is.EqualTo(UnitMetric.EM));
                Assert.That(font.Size.Value, Is.EqualTo(1.2));
            });
        }

        [Test]
        public void WithFontSizeLineHeight_ShouldSucceed()
        {
            var font = HtmlFont.Parse("italic small-caps bold 12px/30px Georgia, serif");
            Assert.Multiple(() => {
                Assert.That(font.Variant, Is.EqualTo(FontVariant.SmallCaps));
                Assert.That(font.Style, Is.EqualTo(FontStyle.Italic));
                Assert.That(font.Weight, Is.EqualTo(FontWeight.Bold));
                Assert.That(font.Family, Is.EqualTo("Georgia"));
                Assert.That(font.Size.Metric, Is.EqualTo(UnitMetric.Pixel));
                Assert.That(font.Size.Value, Is.EqualTo(12));
                Assert.That(font.LineHeight.Metric, Is.EqualTo(UnitMetric.Pixel));
                Assert.That(font.LineHeight.Value, Is.EqualTo(30));
            });
        }
    }
}
