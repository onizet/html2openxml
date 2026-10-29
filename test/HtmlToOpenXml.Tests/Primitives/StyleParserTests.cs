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
        [TestCase("text-decoration &#x3A; underline &#x3B;color :red")]
        [TestCase("text-decoration &colon; underline &semi;color :red")]
        [TestCase("text-decoration &COLON; underline &SEMI;color :red")]
        public void ParseStyle_ShouldSucceed(string htmlStyle)
        {
            var styles = HtmlAttributeCollection.ParseStyle(htmlStyle);
            using (Assert.EnterMultipleScope())
            {
                Assert.That(styles["text-decoration"].ToString(), Is.EqualTo("underline"));
                Assert.That(styles["color"].ToString(), Is.EqualTo("red"));
            }
        }

        [Test(Description = "Parser should consider the last occurence of a style")]
        public void DuplicateStyle_ReturnsLatter()
        {
            var styleAttributes = HtmlAttributeCollection.ParseStyle("color:red;color:BLUE");
            Assert.That(styleAttributes["color"].ToString(), Is.EqualTo("BLUE"));
        }

        [TestCase("color;color;")]
        [TestCase(":;")]
        [TestCase("color:;")]
        public void InvalidStyle_ShouldBeEmpty(string htmlStyle)
        {
            var styles = HtmlAttributeCollection.ParseStyle(htmlStyle);
            Assert.That(styles.IsEmpty, Is.True);
            Assert.That(styles.TryGetValue("color", out var _), Is.False);
        }

        [Test]
        public void WithUnknownEntity_ContinueOnNextKey()
        {
            var styles = HtmlAttributeCollection.ParseStyle("color:red&green;text-decoration:underline");
            Assert.That(styles.IsEmpty, Is.False);
            Assert.That(styles["text-decoration"].ToString(), Is.EqualTo("underline"));
        }

        [Test]
        public void WithMultipleTextDecoration_ReturnsAllValues()
        {
            var styles = HtmlAttributeCollection.ParseStyle("text-decoration:underline dotted wavy");
            var decorations = styles.GetTextDecorations();
            Assert.That(decorations, Is.EqualTo(TextDecoration.Underline| TextDecoration.Dotted| TextDecoration.Wave));
        }

        [Test]
        public void WithEmbedSemicolons_ReturnsLegitimateUrl()
        {
            var styles = HtmlAttributeCollection.ParseStyle("background:url(data:image/svg+xml;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==)");
            var exist = styles.TryGetValue("background", out var imageUrl);
            Assert.That(exist, Is.True);
            if (exist)
            {
                Assert.That(imageUrl.ToString(), Does.StartWith("url("));
                Assert.That(imageUrl.ToString(), Does.EndWith(")"));
                var dataUri = imageUrl.Slice(4, imageUrl.Length - 4 - 1).ToString();
                var success = IO.DataUri.TryCreateImage(dataUri, out var _);
                Assert.That(success, Is.True);
            }
        }

        [Test(Description = "Encoded ':' and ';' characters are valid")]
        public void EncodedStyle_ShouldSucceed()
        {
            var styleAttributes = HtmlAttributeCollection.ParseStyle("text-decoration&#58;underline&#59;color:red");
            Assert.That(styleAttributes["text-decoration"].ToString(), Is.EqualTo("underline"));
            Assert.That(styleAttributes["color"].ToString(), Is.EqualTo("red"));
        }

        [Test(Description = "Key style with no value should be ignored")]
        public void EmptyStyle_ShouldBeIgnored()
        {
            var styleAttributes = HtmlAttributeCollection.ParseStyle("text-decoration;color:red");
            Assert.That(styleAttributes.TryGetValue("text-decoration", out var _), Is.False);
            Assert.That(styleAttributes["color"].ToString(), Is.EqualTo("red"));
        }
    }
}
