using NUnit.Framework;

namespace HtmlToOpenXml.Tests.Primitives
{
    /// <summary>
    /// Tests Html color style attribute.
    /// </summary>
    [TestFixture]
    class UnitTests
    {
        [TestCase("auto", 0, UnitMetric.Auto)]
        [TestCase("AUTO", 0, UnitMetric.Auto, Description = "Should be case insensitive")]
        [TestCase("5%", 5, UnitMetric.Percent)]
        [TestCase(" 12 px", 12, UnitMetric.Pixel)]
        [TestCase(" 12 ", 12, UnitMetric.Unitless)]
        [TestCase("9", 9, UnitMetric.Unitless)]
        public void ParseHtmlUnit_ShouldSucceed(string str, double value, UnitMetric metric)
        {
            var unit = Unit.Parse(str);

            Assert.Multiple(() => {
                Assert.That(unit.IsValid, Is.True);
                Assert.That(unit.Metric, Is.EqualTo(metric));
                Assert.That(unit.Value, Is.EqualTo(value));
            });
        }

        [TestCase("    ")]
        [TestCase("12zz")]
        [TestCase("zz")]
        [TestCase("%")]
        public void ParseInvalidHtmlColor_ReturnsEmpty(string str)
        {
            var unit = Unit.Parse(str);
            Assert.That(unit.IsValid, Is.False);
        }
    }
}
