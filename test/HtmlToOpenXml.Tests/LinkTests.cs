using System;
using System.Linq;
using NUnit.Framework;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests hyperlink.
    /// </summary>
    [TestFixture]
    public class LinkTests : HtmlConverterTestBase
    {
        [Test]
        public void ParseLink()
        {
            var elements = converter.Parse(@"<a href=""www.site.com"" title=""Test Tooltip"">Test Caption</a>");
            Assert.That(elements.Count, Is.EqualTo(1));
            Assert.Multiple(() => {
                Assert.That(elements[0], Is.TypeOf(typeof(Paragraph)));
                Assert.That(elements[0].FirstChild, Is.TypeOf(typeof(Hyperlink)));
                Assert.That(elements[0].FirstChild.FirstChild, Is.TypeOf(typeof(Run)));
                Assert.That(elements[0].InnerText, Is.EqualTo("Test Caption"));
            });

            var hyperlink = (Hyperlink) elements[0].FirstChild;
            Assert.IsNotNull(hyperlink.Tooltip);
            Assert.That(hyperlink.Tooltip.Value, Is.EqualTo("Test Tooltip"));

            Assert.IsNotNull(hyperlink.Id);
            Assert.That(hyperlink.History.Value, Is.EqualTo(true));
            Assert.That(mainPart.HyperlinkRelationships.Count(), Is.GreaterThan(0));

            var extLink = mainPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id);
            Assert.IsNotNull(extLink);
            Assert.That(extLink.IsExternal, Is.EqualTo(true));
            Assert.That(extLink.Uri.AbsoluteUri, Is.EqualTo("http://www.site.com/"));
        }

        [TestCase(@"<a href=""javascript:alert()"">Js</a>")]
        [TestCase(@"<a href=""site.com"">Unknow site</a>")]
        public void ParseInvalidLink (string html)
        {
            // invalid link leads to simple Run with no link

            var elements = converter.Parse(html);
            Assert.That(elements.Count, Is.EqualTo(1));
            Assert.Multiple(() => {
                Assert.That(elements[0], Is.TypeOf(typeof(Paragraph)));
                Assert.That(elements[0].FirstChild, Is.TypeOf(typeof(Run)));
                Assert.That(elements[0].FirstChild.FirstChild, Is.TypeOf(typeof(Text)));
            });
        }

        [Test]
        public void ParseTextImageLink ()
        {
            var elements = converter.Parse(@"<a href=""www.site.com""><img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg=="" alt=""Red dot"" /> Test Caption</a>");
            Assert.That(elements[0].FirstChild, Is.TypeOf(typeof(Hyperlink)));

            var hyperlink = (Hyperlink) elements[0].FirstChild;
            Assert.That(hyperlink.ChildElements.Count, Is.EqualTo(2));
            Assert.That(hyperlink.FirstChild, Is.TypeOf(typeof(Run)));
            Assert.That(hyperlink.FirstChild.FirstChild, Is.TypeOf(typeof(Drawing)));
            Assert.That(hyperlink.LastChild.InnerText, Is.EqualTo(" Test Caption"));
        }

        [Test]
        public void ParseAnchorLink ()
        {
            var elements = converter.Parse(@"<a href=""#anchor1"">Anchor1</a>");
            Assert.That(elements.Count, Is.EqualTo(1));
            Assert.That(elements[0], Is.TypeOf(typeof(Paragraph)));
            Assert.That(elements[0].FirstChild, Is.TypeOf(typeof(Hyperlink)));

            var hyperlink = (Hyperlink) elements[0].FirstChild;
            Assert.IsNull(hyperlink.Id);
            Assert.True(hyperlink.Anchor == "anchor1");

            converter.ExcludeLinkAnchor = true;

            // _top is always present and bypass the previous rule
            elements = converter.Parse(@"<a href=""#_top"">Anchor2</a>");
            hyperlink = (Hyperlink) elements[0].FirstChild;
            Assert.True(hyperlink.Anchor == "_top");

            // this should generate a Run and not an Hyperlink
            elements = converter.Parse(@"<a href=""#_anchor3"">Anchor3</a>");
            Assert.That(elements[0].FirstChild, Is.TypeOf(typeof(Run)));

            converter.ExcludeLinkAnchor = false;
        }
    }
}