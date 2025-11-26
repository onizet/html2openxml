using NUnit.Framework;

namespace HtmlToOpenXml.Tests.Primitives
{
    /// <summary>
    /// Tests parsing the `style` attribute.
    /// </summary>
    [TestFixture]
    public class StyleParserTests
    {
        [TestCase("text-decoration:underline; color: red ")]
        [TestCase("text-decoration &#58; underline &#59;color :red")]
        public void ParseStyle_ShouldSucceed(string htmlStyle)
        {
            var styles = HtmlAttributeCollection.ParseStyle(htmlStyle);
            Assert.Multiple(() => {
                Assert.That(styles.HasKeyEqualsTo("text-decoration", "underline"), Is.True);
                Assert.That(styles.HasKeyEqualsTo("color", "red"), Is.True);
                Assert.That(styles["color"].ToString(), Is.EqualTo("red"));
            });
        }

        [Test(Description = "Parser should consider the last occurence of a style")]
        public void DuplicateStyle_ReturnsLatter()
        {
            var styleAttributes = HtmlAttributeCollection.ParseStyle("color:red;color:BLUE");
            Assert.That(styleAttributes.HasKeyEqualsTo("color", "blue"), Is.True);
        }

        [TestCase("color;color;")]
        [TestCase(":;")]
        [TestCase("color:;")]
        public void InvalidStyle_ShouldBeEmpty(string htmlStyle)
        {
            var styles = HtmlAttributeCollection.ParseStyle(htmlStyle);
            Assert.That(styles.IsEmpty, Is.True);
            Assert.That(styles.ContainsKey("color"), Is.False);
        }

        [Test]
        public void WithMultipleTextDecoration_ReturnsAllValues()
        {
            var styles = HtmlAttributeCollection.ParseStyle("text-decoration:underline dotted wavy");
            var decorations = styles.GetTextDecorations("text-decoration");
            Assert.That(decorations, Is.EquivalentTo([TextDecoration.Underline, TextDecoration.Dotted, TextDecoration.Wave]));
        }
    }
}
