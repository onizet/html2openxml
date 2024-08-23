using NUnit.Framework;

namespace HtmlToOpenXml.Tests.Primitives
{
    /// <summary>
    /// Tests Html margin style attribute.
    /// </summary>
    [TestFixture]
    public class MarginTests
    {
        [TestCase("25px 50px 75px 100px", 25, 50, 75, 100)]
        [TestCase("25px 50px 75px", 25, 50, 75, 50)]
        [TestCase("25px 50px", 25, 50, 25, 50)]
        [TestCase("25px", 25, 25, 25, 25)]
        public void ParseHtmlString_ShouldSucceed (string html, int top, int right, int bottom, int left)
        {
            var margin = Margin.Parse(html);

            Assert.Multiple(() => {
                Assert.That(margin.IsValid, Is.EqualTo(true));
                Assert.That(margin.Top.ValueInPx, Is.EqualTo(top));
                Assert.That(margin.Right.ValueInPx, Is.EqualTo(right));
                Assert.That(margin.Bottom.ValueInPx, Is.EqualTo(bottom));
                Assert.That(margin.Left.ValueInPx, Is.EqualTo(left));
            });
        }

        [Test]
        public void ParseWithFloat_ShouldSucceed ()
        {
            var margin = Margin.Parse("0 50% 9.5pt .00001pt");

            Assert.Multiple(() => {
                Assert.That(margin.IsValid, Is.EqualTo(true));

                Assert.That(margin.Top.Value, Is.EqualTo(0));
                Assert.That(margin.Top.Type, Is.EqualTo(UnitMetric.Pixel));

                Assert.That(margin.Right.Value, Is.EqualTo(50));
                Assert.That(margin.Right.Type, Is.EqualTo(UnitMetric.Percent));

                Assert.That(margin.Bottom.Value, Is.EqualTo(9.5));
                Assert.That(margin.Bottom.Type, Is.EqualTo(UnitMetric.Point));
                Assert.That(margin.Bottom.ValueInPoint, Is.EqualTo(9.5));
                //size are half-point font size (OpenXml relies mostly on long value, not on float)
                Assert.That(Math.Round(margin.Bottom.ValueInPoint * 2).ToString(), Is.EqualTo("19"));

                Assert.That(margin.Left.Value, Is.EqualTo(.00001));
                Assert.That(margin.Left.Type, Is.EqualTo(UnitMetric.Point));
                // but due to conversion: 0 (OpenXml relies mostly on long value, not on float)
                Assert.That(Math.Round(margin.Left.ValueInPoint * 2).ToString(), Is.EqualTo("0"));
            });
        }

        [Test]
        public void ParseWithAuto_ShouldSucceed ()
        {
            var margin = Margin.Parse("0 auto");

            Assert.Multiple(() => {
                Assert.That(margin.IsValid, Is.EqualTo(true));

                Assert.That(margin.Top.Value, Is.EqualTo(0));
                Assert.That(margin.Top.Type, Is.EqualTo(UnitMetric.Pixel));

                Assert.That(margin.Bottom.Value, Is.EqualTo(0));
                Assert.That(margin.Bottom.Type, Is.EqualTo(UnitMetric.Pixel));

                Assert.That(margin.Left.Type, Is.EqualTo(UnitMetric.Auto));
                Assert.That(margin.Right.Type, Is.EqualTo(UnitMetric.Auto));
            });
        }
    }
}
