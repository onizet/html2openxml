using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests parser to control the various whitespacing handling.
    /// </summary>
    /// <sealso cref="https://developer.mozilla.org/en-US/docs/Web/API/Document_Object_Model/Whitespace" />
    [TestFixture]
    public class WhitespaceTests : HtmlConverterTestBase
    {
        [Test]
        public void ParseConsecutiveRuns ()
        {
            // the new line should generate a space between "bold" and "text"
            var elements = converter.Parse("<span>This is a <b>bold\n</b>text</span>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0].ChildElements, Is.All.TypeOf<Run>());
            Assert.That(elements[0].InnerText, Is.EqualTo("This is a bold text"));
        }

        [Test]
        public void ParseConsecutiveParagraphs ()
        {
            var elements = converter.Parse("<div>Hello</div><div>World</div>");
            Assert.That(elements, Has.Count.EqualTo(2));
            Assert.That(elements, Is.All.TypeOf<Paragraph>());
            Assert.That(elements[0].InnerText, Is.EqualTo("Hello"));
            Assert.That(elements[1].InnerText, Is.EqualTo("World"));
        }

        [Test]
        public void ParseInlineElements ()
        {
            var elements = converter.Parse(@"<h1>   Hello
        <span> World!</span>   </h1>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0].InnerText, Is.EqualTo("Hello World!"));
        }

        [TestCase("h1", false)]
        [TestCase("pre", true)]
        [TestCase("span", false)]
        [TestCase("p", false)]
        [TestCase("a", false)]
        public void ParseWhitespace(string tagName, bool expectWhitespaces)
        {
            var elements = converter.Parse($"<{tagName}>      Hello      World!     </{tagName}>");
            Assert.That(elements, Has.Count.EqualTo(1));

            string expectedText = expectWhitespaces? "      Hello      World!     " : "Hello World!";
            Assert.That(elements[0].InnerText, Is.EqualTo(expectedText));
        }

        [Test(Description = "When the anchor is prefixed by an image, the initial whitespace is collapsed")]
        public void ParseWhitespaceAnchorWithImg()
        {
            var elements = converter.Parse(@"<a><img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==""/>     Hello      World!     </a>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0].InnerText, Is.EqualTo(" Hello World!"));
        }

        [Test(Description = "`nbsp` entities should not be collapsed")]
        public void ParseNonBreakingSpace()
        {
            var elements = converter.Parse("<h1>&nbsp;&nbsp; Hello      World!     </h1>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0].InnerText, Is.EqualTo("   Hello World!"));
        }
    }
}