using NUnit.Framework;

namespace HtmlToOpenXml.Tests
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
        public void Parse (string html, int top, int right, int bottom, int left)
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
        public void ParseFloat ()
        {
            var margin = Margin.Parse("0 50% 1em .00001pt");

            Assert.Multiple(() => {
                Assert.That(margin.IsValid, Is.EqualTo(true));

                Assert.That(margin.Top.Value, Is.EqualTo(0));
                Assert.That(margin.Top.Type, Is.EqualTo(UnitMetric.Pixel));

                Assert.That(margin.Right.Value, Is.EqualTo(50));
                Assert.That(margin.Right.Type, Is.EqualTo(UnitMetric.Percent));

                Assert.That(margin.Bottom.Value, Is.EqualTo(1));
                Assert.That(margin.Bottom.Type, Is.EqualTo(UnitMetric.EM));

                Assert.That(margin.Left.Value, Is.EqualTo(.00001));
                Assert.That(margin.Left.Type, Is.EqualTo(UnitMetric.Point));
                // but due to conversion: 0 (OpenXml relies mostly on long value, not on float)
                Assert.That(margin.Left.ValueInPoint, Is.EqualTo(0));
            });
        }
    }
}
