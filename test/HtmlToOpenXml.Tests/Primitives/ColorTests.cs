using NUnit.Framework;

namespace HtmlToOpenXml.Tests.Primitives
{
    /// <summary>
    /// Tests Html color style attribute.
    /// </summary>
    [TestFixture]
    public class ColorTests
    {
        [TestCase("", 0, 0, 0, 0d)]
        [TestCase("#F00", 255, 0, 0, 1d)]
        [TestCase("#00FFFF", 0, 255, 255, 1d)]
        [TestCase("red", 255, 0, 0, 1d)]
        [TestCase("rgb(106, 90, 205)", 106, 90, 205, 1d)]
        [TestCase("rgba(106, 90, 205, 0.6)", 106, 90, 205, 0.6d)]
        [TestCase("hsl(248, 53%, 58%)", 106, 91, 205, 1)]
        [TestCase("hsla(9, 100%, 64%, 0.6)", 255, 99, 71, 0.6d)]
        [TestCase("hsl(0, 100%, 50%)", 255, 0, 0, 1)]
        // Percentage not respected that should be maxed out
        [TestCase("hsl(0, 200%, 150%)", 255, 255, 255, 1)]
        // Failure that leads to empty
        [TestCase("rgba(1.06, 90, 205, 0.6)", 0, 0, 0, 0.0d)]
        [TestCase("rgba(a, r, g, b)", 0, 0, 0, 0.0d)]
        public void ParseHtmlColor_ShouldSucceed(string htmlColor, byte red, byte green, byte blue, double alpha)
        {
            var color = HtmlColor.Parse(htmlColor);

            Assert.Multiple(() => {
                Assert.That(color.R, Is.EqualTo(red));
                Assert.That(color.B, Is.EqualTo(blue));
                Assert.That(color.G, Is.EqualTo(green));
                Assert.That(color.A, Is.EqualTo(alpha));
            });
        }

        [TestCase(255, 0, 0, 0, ExpectedResult = "FF0000")]
        public string ArgColor_ToHex_ShouldSucceed(byte red, byte green, byte blue, double alpha)
        {
            var color = HtmlColor.FromArgb(alpha, red, green, blue);
            return color.ToHexString();
        }

        [TestCase(0, 248, 0.53, 0.58, ExpectedResult = "6A5BCD")]
        public string HslColor_ToHex_ShouldSucceed(double alpha, double hue, double saturation, double luminosity)
        {
            var color = HtmlColor.FromHsl(alpha, hue, saturation, luminosity);
            return color.ToHexString();
        }
    }
}
