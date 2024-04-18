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
        [Test]
        public void ParseNumberingHeading()
        {
            // the inner html shouldn't be interpreted
            var elements = converter.Parse(@"
                <h1>1. Heading 1<h1>
                <h2>1.1  Heading 1.1<h2><!-- double space after number -->
                <h2>1.2 Heading 1.2<h2><!-- tab after number -->
            ");

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .Where(abs => abs.AbstractNumDefinitionName.Val == NumberingExpression.HeadingNumberingName)
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            var inst = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => i.AbstractNumId.Val == absNum.AbstractNumberId)
                .FirstOrDefault();
            Assert.That(inst, Is.Not.Null);
            Assert.That(inst.NumberID?.Value, Is.Not.Null);

            var paragraphs = elements.Cast<Paragraph>();
            Assert.Multiple(() =>
            {
                Assert.That(paragraphs.Count(), Is.EqualTo(3));
                Assert.That(paragraphs.Select(p => p.InnerText),
                    Has.All.StartsWith("Heading"),
                    "Number and whitespaces are trimmed");
                Assert.That(paragraphs.Select(e =>
                     e.ParagraphProperties.NumberingProperties?.NumberingId?.Val?.Value),
                     Has.All.EqualTo(inst.NumberID.Value),
                     "All paragraphs are linked to the same list instance");
                Assert.That(paragraphs.First().ParagraphProperties.NumberingProperties?.NumberingLevelReference?.Val?.Value,
                    Is.EqualTo(0),
                    "First paragraph stands on level 0");
                Assert.That(paragraphs.Skip(1).Select(e => 
                    e.ParagraphProperties.NumberingProperties?.NumberingLevelReference?.Val?.Value),
                    Has.All.EqualTo(1),
                    "All paragraphs stand on level 1");
            });

        }
    }
}