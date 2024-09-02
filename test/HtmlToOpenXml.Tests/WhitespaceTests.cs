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
        public void ConsecutivePhrasing_ReturnsOneParagraphWithMulitpleRuns ()
        {
            // the new line should generate a space between "bold" and "text"
            var elements = converter.Parse("<span>This is a <b>bold\n</b>text</span>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].ChildElements, Is.All.TypeOf<Run>());
            Assert.That(elements[0].Elements<Run>().Count(), Is.GreaterThan(1));
            Assert.That(elements[0].InnerText, Is.EqualTo("This is a bold text"));
        }

        [Test]
        public void ConsecutiveDivs_ReturnsMultipleParagraphs ()
        {
            var elements = converter.Parse("<div>Hello</div><div>World</div>");
            Assert.That(elements, Has.Count.EqualTo(2));
            Assert.That(elements, Is.All.TypeOf<Paragraph>());
            Assert.That(elements[0].InnerText, Is.EqualTo("Hello"));
            Assert.That(elements[1].InnerText, Is.EqualTo("World"));
        }

        [TestCase("<h1>   Hello\r\n<span> World!</span>   </h1>")]
        [TestCase("<span>   Hello \r\n World!   </span>")]
        [TestCase("<span>   Hello\r\n\r\nWorld!   </span>")]
        public void Multiline_ReturnsCollapsedText (string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].InnerText, Is.EqualTo("Hello World!"));
        }

        [TestCase("h1")]
        [TestCase("span")]
        [TestCase("p")]
        [TestCase("a")]
        public void HtmlTag_ReturnsTrimmedSpaces(string tagName)
        {
            var elements = converter.Parse($"<{tagName}>      Hello      World!     </{tagName}>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].InnerText, Is.EqualTo("Hello World!"));
        }

        [Test]
        public void PreTag_ReturnsPreservedSpaces()
        {
            var elements = converter.Parse($"<pre>      Hello      World!     </pre>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].InnerText, Is.EqualTo("      Hello      World!     "));
        }

        [Test(Description = "When the anchor is prefixed by an image, the initial whitespace is collapsed")]
        public void AnchorWithImgThenText_ReturnsCollapsedStartingWhitespace()
        {
            var elements = converter.Parse(@"<a><img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==""/>     Hello      World!     </a>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].InnerText, Is.EqualTo(" Hello World!"));
        }

        [Test(Description = "`nbsp` entities should not be collapsed")]
        public void NonBreakingSpaceEntities_ReturnsPreservedWhitespace()
        {
            var elements = converter.Parse("<h1>&nbsp;&nbsp; Hello      World!     </h1>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].InnerText, Is.EqualTo("   Hello World!"));
        }

        [Test(Description = "Consecutive runs separated by a break should not prefix the 2nd line with a space")]
        public void ConsecutivePhrasingWithBreak_ReturnsSecondLineWithNoSpaces()
        {
            var elements = converter.Parse("<span>Hello<br><span>World</span></span>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].InnerText, Is.EqualTo("HelloWorld"));
            var runs = elements[0].Elements<Run>();
            Assert.That(runs.Count(), Is.EqualTo(3));
            Assert.Multiple(() => {
                Assert.That(runs.ElementAt(1).LastChild, Is.TypeOf<Break>());
                Assert.That(runs.ElementAt(2).FirstChild, Is.TypeOf<Text>());
            });
            Assert.That(((Text)runs.ElementAt(2).FirstChild).Text, Is.EqualTo("World"));
        }
    }
}