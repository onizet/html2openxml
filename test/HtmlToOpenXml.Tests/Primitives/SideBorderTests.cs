using DocumentFormat.OpenXml;
using NUnit.Framework;

namespace HtmlToOpenXml.Tests.Primitives
{
    /// <summary>
    /// Tests Html border style attribute.
    /// </summary>
    [TestFixture]
    public class SideBorderTests
    {
        [TestCase("solid #ff0000", "single", 255, 0, 0)]
        [TestCase("1px dashed rgb(233, 233, 233)", "dashed", 233, 233, 233)]
        [TestCase("thin dotted white", "dotted", 255, 255, 255)]
        public void ParseHtmlBorder_ShouldSucceed(string htmlBorder, string borderStyle, byte red, byte green, byte blue)
        {
            var border = SideBorder.Parse(htmlBorder);

            Assert.Multiple(() => {
                Assert.That(border.IsValid, Is.True);
                Assert.That(((IEnumValue) border.Style).Value, Is.EqualTo(borderStyle));
                Assert.That(border.Color.R, Is.EqualTo(red));
                Assert.That(border.Color.B, Is.EqualTo(blue));
                Assert.That(border.Color.G, Is.EqualTo(green));
            });
        }

        [TestCase("")]
        [TestCase("abc")]
        public void InvalidBorder_ShouldFail(string htmlBorder)
        {
            var border = SideBorder.Parse(htmlBorder);
            Assert.That(border.IsValid, Is.False);
        }

        [Test]
        public void Border_ShouldSucceed()
        {
            var border = SideBorder.Parse("3px solid black");
            Assert.That(border.IsValid, Is.True);
            Assert.That(border.Width.ValueInPx, Is.EqualTo(3));
            Assert.That(border.Width.ValueInPoint, Is.EqualTo(2.25));
            Assert.That(border.Width.ValueInEighthPoint, Is.EqualTo(18));
        }
    }
}
