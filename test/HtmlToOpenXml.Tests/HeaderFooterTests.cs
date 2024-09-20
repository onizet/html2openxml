using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests on <c>ParseHeader</c> and <c>ParseFooter</c> methods.
    /// </summary>
    [TestFixture]
    public class HeaderFooterTests : HtmlConverterTestBase
    {
        [Test]
        public async Task Header_ReturnsHeaderPartLinkedToBody()
        {
            await converter.ParseHeader("<p>Header content</p>");

            var headerPart = mainPart.HeaderParts?.FirstOrDefault();
            Assert.That(headerPart, Is.Not.Null);
            Assert.That(headerPart.Header, Is.Not.Null);

            var sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Empty);
            Assert.That(sectionProperties.Any(s => s.HasChild<HeaderReference>()), Is.True);
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test]
        public async Task Footer_ReturnsFooterPartLinkedToBody()
        {
            await converter.ParseFooter("<p>Footer content</p>");

            var footerPart = mainPart.FooterParts?.FirstOrDefault();
            Assert.That(footerPart, Is.Not.Null);
            Assert.That(footerPart.Footer, Is.Not.Null);

            var sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Empty);
            Assert.That(sectionProperties.Any(s => s.HasChild<FooterReference>()), Is.True);
            AssertThatOpenXmlDocumentIsValid();
        }
    }
}