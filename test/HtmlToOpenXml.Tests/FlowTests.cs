using NUnit.Framework;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests on <c>div</c> and other block elements.
    /// </summary>
    [TestFixture]
    public class FlowTests : HtmlConverterTestBase
    {
        [Test]
        public void ParsePageBreakBefore()
        {
            var elements = converter.Parse(@"Lorem
                <div style='page-break-before:always'>Placeholder</div>
                Ipsum");
            Assert.That(elements, Has.Count.EqualTo(3));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].ChildElements, Has.Count.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(elements[0].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[0].InnerText, Is.EqualTo("Lorem"));
                Assert.That(elements[1].ChildElements, Has.Count.EqualTo(3));
                Assert.That(elements[2].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[2].InnerText, Is.EqualTo("Ipsum"));
                Assert.That(elements[2].ChildElements, Has.Count.EqualTo(1));
            });
            Assert.Multiple(() =>
            {
                Assert.That(elements[1].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[1].ChildElements[0].HasChild<Break>(), Is.True);
                Assert.That(elements[1].ChildElements[1].HasChild<LastRenderedPageBreak>(), Is.True);
                Assert.That(elements[1].ChildElements[2].InnerText, Is.EqualTo("Placeholder"));
                Assert.That(elements[1].InnerText, Is.EqualTo("Placeholder"));
            });
        }

        [Test]
        public void ParsePageBreakAfter()
        {
            var elements = converter.Parse(@"Lorem
                <div style='page-break-after:always'>Placeholder</div>
                Ipsum");
            Assert.That(elements, Has.Count.EqualTo(3));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].ChildElements, Has.Count.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(elements[0].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[0].InnerText, Is.EqualTo("Lorem"));
                Assert.That(elements[1].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[1].InnerText, Is.EqualTo("Placeholder"));
                Assert.That(elements[2].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[2].InnerText, Is.EqualTo("Ipsum"));
            });
            Assert.That(elements[1].LastChild.HasChild<Break>(), Is.True);
            Assert.That(elements[1].LastChild.HasChild<LastRenderedPageBreak>(), Is.False);
        }

        [TestCase("landscape")]
        [TestCase("portrait")]
        public void ParsePageOrientation(string orientation)
        {
            var _ = converter.Parse($@"<body style=""page-orientation:{orientation}""><body>");
            var sectionProperties = mainPart.Document.Body!.GetFirstChild<SectionProperties>();
            Assert.That(sectionProperties, Is.Not.Null);
            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            if (orientation == "landscape")
                Assert.That(pageSize.Width, Is.GreaterThan(pageSize.Height));
            else
                Assert.That(pageSize.Height, Is.GreaterThan(pageSize.Width));
        }
    }
}