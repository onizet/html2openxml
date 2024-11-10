using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml.Expressions;
using NUnit.Framework;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests for heading.
    /// </summary>
    [TestFixture]
    public class HeadingTests : HtmlConverterTestBase
    {
        [TestCase("<h1>1. Heading 1</h1><h2>1.1 Heading Normal Case</h1>")]
        [TestCase("<h1>1. Heading 1</h1><h2>1.1  Heading Double Space</h2>", Description = "Double space after number")]
        [TestCase("<h1>1. Heading 1</h1><h2>1.2&#09;Heading Tab</h2>", Description = "Tab after number")]
        [TestCase("<h1>1. Heading 1</h1><h2>1.3Heading No Space</h2>", Description = "No space after number")]
        public void OrderedPattern_ReturnsNumberingHeading(string html)
        {
            var elements = converter.Parse(html);

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .Where(abs => abs.AbstractNumDefinitionName?.Val == NumberingExpressionBase.HeadingNumberingName)
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            var inst = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => i.AbstractNumId?.Val == absNum.AbstractNumberId)
                .FirstOrDefault();
            Assert.That(inst, Is.Not.Null);
            Assert.That(inst.NumberID?.Value, Is.Not.Null);

            var paragraphs = elements.Cast<Paragraph>();
            Assert.Multiple(() =>
            {
                Assert.That(paragraphs.Count(), Is.EqualTo(2));
                Assert.That(paragraphs.Select(p => p.InnerText),
                    Has.All.StartsWith("Heading"),
                    "Number and whitespaces are trimmed");
                Assert.That(paragraphs.Select(e =>
                     e.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value),
                     Has.All.EqualTo(inst.NumberID.Value),
                     "All paragraphs are linked to the same list instance");
                Assert.That(paragraphs.First().ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value,
                    Is.EqualTo(0),
                    "First paragraph stands on level 0");
                Assert.That(paragraphs.Skip(1).Select(e => 
                    e.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value),
                    Has.All.EqualTo(1),
                    "All paragraphs stand on level 1");
            });
        }

        [TestCase("<h1>1. Heading 1</h1><h2>1.1 Heading Normal Case</h1>")]
        [TestCase("<h1>1. Heading 1</h1><h2>1.1  Heading Double Space</h2>", Description = "Double space after number")]
        [TestCase("<h1>1. Heading 1</h1><h2>1.2&#09;Heading Tab</h2>", Description = "Tab after number")]
        [TestCase("<h1>1. Heading 1</h1><h2>1.3Heading No Space</h2>", Description = "No space after number")]
        public void OrderedPattern_DisableNumberingSupports_ReturnsSimpleHeading(string html)
        {
            converter.SupportsHeadingNumbering = false;
            var elements = converter.Parse(html);

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .Where(abs => abs.AbstractNumDefinitionName?.Val == NumberingExpressionBase.HeadingNumberingName)
                .SingleOrDefault();
            Assert.That(absNum, Is.Null);

            var paragraphs = elements.Cast<Paragraph>();
            Assert.Multiple(() =>
            {
                Assert.That(paragraphs.Count(), Is.EqualTo(2));
                Assert.That(paragraphs.First().InnerText, Is.EqualTo("1. Heading 1"));
                Assert.That(paragraphs.First().ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val,
                    Is.Null,
                    "First paragraph is not a numbering");
            });
        }

        [Test]
        public void MaxLevel_ShouldBeIgnored()
        {
            const int maxLevel = 6;
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i <= maxLevel; i++)
                sb.AppendFormat("<h{0}>Heading {0}</h{0}>", i+1);

            var elements = converter.Parse(sb.ToString());

            Assert.Multiple(() =>
            {
                Assert.That(elements.Count(), Is.EqualTo(maxLevel + 1));
                Assert.That(elements, Has.All.TypeOf<Paragraph>());
            });

            Assert.Multiple(() =>
            {
                Assert.That(elements.Take(maxLevel).Select(p => p.GetFirstChild<ParagraphProperties>()?.ParagraphStyleId?.Val?.Value),
                            Has.All.StartsWith("Heading"));
                Assert.That(elements.Last().GetFirstChild<ParagraphProperties>()?.ParagraphStyleId,
                    Is.Null, $"Only {maxLevel+1} levels of heading supported");
            });
        }
    }
}