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
            var border = SideBorder.Parse(htmlBorder.AsSpan());

            using (Assert.EnterMultipleScope())
            {
                Assert.That(border.IsValid, Is.True);
                Assert.That(((IEnumValue) border.Style).Value, Is.EqualTo(borderStyle));
                Assert.That(border.Color.R, Is.EqualTo(red));
                Assert.That(border.Color.B, Is.EqualTo(blue));
                Assert.That(border.Color.G, Is.EqualTo(green));
            }
        }

        [TestCase("")]
        [TestCase("abc")]
        public void InvalidBorder_ShouldFail(string htmlBorder)
        {
            var border = SideBorder.Parse(htmlBorder.AsSpan());
            Assert.That(border.IsValid, Is.False);
        }

        [Test]
        public void Border_ShouldSucceed()
        {
            var border = SideBorder.Parse("3px solid black".AsSpan());
            Assert.That(border.IsValid, Is.True);
            Assert.That(border.Width.ValueInPx, Is.EqualTo(3));
            Assert.That(border.Width.ValueInPoint, Is.EqualTo(2.25));
            Assert.That(border.Width.ValueInEighthPoint, Is.EqualTo(18));
        }

        [TestCase("1pt", new[] { 1d })]
        [TestCase("1pt 2pt", new[] { 1d, 2d })]
        [TestCase("1pt 2pt 3pt", new[] { 1d, 2d, 3d })]
        [TestCase("1pt 2pt 3pt 4pt", new[] { 1d, 2d, 3d, 4d })]
        public void ParseMultipleWidth_ShouldParseCssWidthValues(string value, double[] expected)
        {
            var widths = SideBorder.ParseMultipleWidth(value.AsSpan());

            Assert.That(widths, Has.Length.EqualTo(expected.Length));

            using (Assert.EnterMultipleScope())
            {
                for (int i = 0; i < expected.Length; i++)
                {
                    Assert.That(widths[i].ValueInPoint, Is.EqualTo(expected[i]).Within(0.01));
                }
            }
        }

        [TestCase("border-width: 1pt; border-style: solid;", 1, 1, 1, 1)]
        [TestCase("border-width: 1pt 2pt; border-style: solid;", 1, 2, 1, 2)]
        [TestCase("border-width: 1pt 2pt 3pt; border-style: solid;", 1, 2, 3, 2)]
        [TestCase("border-width: 1pt 2pt 3pt 4pt; border-style: solid;", 1, 2, 3, 4)]
        public void GetBorders_ShouldExpandBorderWidth(string borderWidth, double top, double right, double bottom, double left)
        {
            var attributes = HtmlAttributeCollection.ParseStyle(borderWidth);
            var border = attributes.GetBorders();
            using (Assert.EnterMultipleScope())
            {
                Assert.That(border.Top.Width.ValueInPoint, Is.EqualTo(top));
                Assert.That(border.Right.Width.ValueInPoint, Is.EqualTo(right));
                Assert.That(border.Bottom.Width.ValueInPoint, Is.EqualTo(bottom));
                Assert.That(border.Left.Width.ValueInPoint, Is.EqualTo(left));
            }
        }

        [TestCase("border-color: red; border-style: solid;", "red", "red", "red", "red")]
        [TestCase("border-color: red blue; border-style: solid;", "red", "blue", "red", "blue")]
        [TestCase("border-color: red blue green; border-style: solid;", "red", "blue", "green", "blue")]
        [TestCase("border-color: red blue green yellow; border-style: solid;", "red", "blue", "green", "yellow")]
        public void GetBorders_ShouldExpandBorderColor(string borderColor, string top, string right, string bottom, string left)
        {
            var attributes = HtmlAttributeCollection.ParseStyle(borderColor);
            var border = attributes.GetBorders();
            using (Assert.EnterMultipleScope())
            {
                Assert.That(border.Top.Color.ToHexString(), Is.EqualTo(HtmlColor.Parse(top).ToHexString()));
                Assert.That(border.Right.Color.ToHexString(), Is.EqualTo(HtmlColor.Parse(right).ToHexString()));
                Assert.That(border.Bottom.Color.ToHexString(), Is.EqualTo(HtmlColor.Parse(bottom).ToHexString()));
                Assert.That(border.Left.Color.ToHexString(), Is.EqualTo(HtmlColor.Parse(left).ToHexString()));
            }
        }

        [TestCase("border-style: solid;", "single", "single", "single", "single")]
        [TestCase("border-style: solid dashed;", "single", "dashed", "single", "dashed")]
        [TestCase("border-style: solid dashed dotted;", "single", "dashed", "dotted", "dashed")]
        [TestCase("border-style: solid dashed dotted double;", "single", "dashed", "dotted", "double")]
        public void GetBorders_ShouldExpandBorderStyle(string borderStyle, string top, string right, string bottom, string left)
        {
            var attributes = HtmlAttributeCollection.ParseStyle(borderStyle);
            var border = attributes.GetBorders();

            using (Assert.EnterMultipleScope())
            {
                Assert.That(((IEnumValue)border.Top.Style).Value, Is.EqualTo(top));
                Assert.That(((IEnumValue)border.Right.Style).Value, Is.EqualTo(right));
                Assert.That(((IEnumValue)border.Bottom.Style).Value, Is.EqualTo(bottom));
                Assert.That(((IEnumValue)border.Left.Style).Value, Is.EqualTo(left));
            }
        }
    }
}
