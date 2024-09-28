using NUnit.Framework;

namespace HtmlToOpenXml.Tests.Primitives
{
    /// <summary>
    /// Tests Html color style attribute.
    /// </summary>
    [TestFixture]
    public class StyleParserTests
    {
        [TestCase("text-decoration:underline; color: red ")]
        [TestCase("text-decoration&#58;underline&#59;color:red")]
        public void ParseStyle_ShouldSucceed(string htmlStyle)
        {
            var styles = HtmlAttributeCollection.ParseStyle(htmlStyle);
            Assert.Multiple(() => {
                Assert.That(styles["text-decoration"], Is.EqualTo("underline"));
                Assert.That(styles["color"], Is.EqualTo("red"));
            });
        }

        [Test(Description = "Parser should consider the last occurence of a style")]
        public void DuplicateStyle_ReturnsLatter()
        {
            var styleAttributes = HtmlAttributeCollection.ParseStyle("color:red;color:blue");
            Assert.That(styleAttributes["color"], Is.EqualTo("blue"));
        }
    }
}
