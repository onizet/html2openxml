using NUnit.Framework;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests Horizontal Lines.
    /// </summary>
    [TestFixture]
    public class HrTests : HtmlConverterTestBase
    {
        [Test]
        public void ParseHr ()
        {
            var elements = converter.Parse("<hr>");
            AssertIsHr (elements[0], false);
        }

        [Test]
        public void ParseHrNoSpacing ()
        {
            // this should not generate a particular spacing
            var elements = converter.Parse("<p style='border-top:1px solid black'>Before</p><hr>");
            AssertIsHr (elements[1], false);
        }

        [TestCase("<p style='border-bottom:1px solid black'>Before</p><hr>")]
        [TestCase("<table><tr><td>Cell</td></tr></table><hr>")]
        public void ParseHrWithSpacing (string html)
        {
            var elements = converter.Parse(html);
            AssertIsHr (elements[1], true);
        }

        private void AssertIsHr (OpenXmlCompositeElement hr, bool expectSpacing)
        {
            Assert.That(hr.ChildElements.Count, Is.EqualTo(2));
            var props = hr.GetFirstChild<ParagraphProperties>();
            Assert.IsNotNull(props);

            Assert.That(props.ChildElements.Count, Is.EqualTo(expectSpacing? 2:1));
            Assert.IsNotNull(props.ParagraphBorders);
            Assert.IsNotNull(props.ParagraphBorders.TopBorder);

            if (expectSpacing)
                Assert.IsNotNull(props.SpacingBetweenLines);
        }
    }
}