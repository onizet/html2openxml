using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests <c>ul</c>, <c>ol</c> and <c>li</c>.
    /// </summary>
    [TestFixture]
    public class NumberingTests : HtmlConverterTestBase
    {
        const int maxLevel = 8;


        [Test(Description = "Skip any elements that is not a `li` tag")]
        public void NonLiElement_ShouldBeIgnored()
        {
            var elements = converter.Parse(@"<ol>
                <p>Must be ignored</p>
                <li>Element1</li>
                <li>Element2</li>
            </ol>");
            Assert.That(elements, Has.Count.EqualTo(2));
            Assert.Multiple(() => {
                Assert.That(elements[0], Is.TypeOf(typeof(Paragraph)));
                Assert.That(elements[0].HasChild<Run>(), Is.True);
                Assert.That(elements[0].InnerText, Does.StartWith("Element"));
            });
        }

        [Test(Description = "Two consecutive lists should restart numbering to 1")]
        public async Task ConsecutiveList_ReturnsList_RestartingOrder()
        {
            await converter.ParseBody(@"
                <oL><li>Item 1.1</li></oL>
                <p>placeholder</p>
                <ol><li>Item 2.1</li></ol>");
            var elements = mainPart.Document.Body!.ChildElements;
            Assert.Multiple(() => {
                Assert.That(elements, Has.Count.EqualTo(3));
                Assert.That(elements, Is.All.TypeOf<Paragraph>());
            });

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            var instances = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => i.AbstractNumId?.Val == absNum.AbstractNumberId);
            Assert.That(instances?.Count(), Is.EqualTo(2));

            Paragraph p1 = (Paragraph) elements[0];
            Paragraph p2 = (Paragraph) elements[2];
            Assert.Multiple(() =>
            {
                Assert.That(new[]{p1, p2}.Select(e => 
                    e.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value),
                    Has.All.EqualTo(0),
                    "All paragraphs stand on level 0");
                Assert.That(p1.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value,
                    Is.Not.EqualTo(p2.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value),
                    "Expected two different list instances");
            });
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test]
        public async Task NestedNumberList_ReturnsMultilevelList()
        {
            await converter.ParseBody(
                @"<ol>
                    <li>Item 1
                        <ol><li>Item 1.1</li></ol>
                    </li>
                    <li>Item 2</li>
                </ol>");
            var elements = mainPart.Document.Body!.ChildElements;
            Assert.Multiple(() => {
                Assert.That(elements, Has.Count.EqualTo(3));
                Assert.That(elements, Is.All.TypeOf<Paragraph>());
                Assert.That(elements[1].InnerText, Is.EqualTo("Item 1.1"));
                Assert.That(mainPart.NumberingDefinitionsPart?.Numbering, Is.Not.Null);
            });

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            // assert numbering template definition
            Assert.Multiple(() =>
            {
                // this is not a real expected constant values but something defined internally in ListExpression
                Assert.That(absNum.AbstractNumDefinitionName?.Val?.Value, Is.EqualTo("decimal"));
                Assert.That(absNum.MultiLevelType?.Val?.InnerText, Is.AnyOf("hybridMultilevel", "multilevel"));
                Assert.That(absNum.Elements<Level>().Count(), Is.AtLeast(2), "At least 2 level registred");
                Assert.That(absNum.GetFirstChild<Level>()?.NumberingFormat?.Val?.Value, Is.EqualTo(NumberFormatValues.Decimal));
            });

            var inst = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => i.AbstractNumId?.Val == absNum.AbstractNumberId)
                .SingleOrDefault();
            Assert.That(inst, Is.Not.Null);
            Assert.That(inst.NumberID?.Value, Is.Not.Null);

            Paragraph p1 = (Paragraph) elements[0];
            Paragraph p1_1 = (Paragraph) elements[1];
            Paragraph p2 = (Paragraph) elements[2];
            // assert paragraphs linked to numbering instance
            Assert.Multiple(() =>
            {
                Assert.That(elements.Cast<Paragraph>().Select(e => 
                    e.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value),
                    Has.All.EqualTo(inst.NumberID.Value),
                    "All paragraphs are linked to the same list instance");
                Assert.That(p1.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value, Is.EqualTo(0));
                Assert.That(p1_1.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value, Is.EqualTo(1));
                Assert.That(p2.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value, Is.EqualTo(0));
            });
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test(Description = "Empty list should not be registred")]
        public void EmptyList_ShouldBeIgnored()
        {
            var elements = converter.Parse("<ol></ol>");
            Assert.That(elements, Is.Empty);
            var numbering = mainPart.NumberingDefinitionsPart?.Numbering;
            if (numbering != null)
            {
                Assert.Multiple(() =>
                {
                    Assert.That(numbering?.Elements<AbstractNum>(), Is.Empty);
                    Assert.That(numbering?.Elements<NumberingInstance>(), Is.Empty);
                });
            }
        }

        [Test(Description = "Empty list item should not be registred")]
        public void EmptyLiElement_ShouldBeIgnored()
        {
            var elements = converter.Parse(@"<ul>
                <li>not empty</li>
                <li></li>
                </ul>");
            Assert.That(elements, Has.Count.EqualTo(1));
        }

        [Test(Description = "Increment instanceId based on existing lists")]
        public void WithExistingNumbering_ReturnsUniqueInstanceId()
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithNumbering.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;
            var numbering = mainPart.NumberingDefinitionsPart?.Numbering;
            Assert.That(numbering, Is.Not.Null);
            var instances = numbering.Elements<NumberingInstance>();
            var beforeMaxInstanceId = instances.MaxBy(i => i.NumberID?.Value)!.NumberID!.Value;
            var beforeInstanceCount = instances.Count();
            Assert.That(beforeInstanceCount, Is.GreaterThan(0));

            HtmlConverter converter = new(mainPart);

            var elements = converter.Parse("<ul><li>Item 1</li></ul>");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(instances.Count(), 
                Is.GreaterThan(beforeInstanceCount),
                "New list instance is appended to existing instances");
            var afterMaxInstanceId = instances.MaxBy(i => i.NumberID?.Value)!.NumberID!.Value;
            Assert.That(afterMaxInstanceId, Is.EqualTo(beforeMaxInstanceId + 1),
                "The new list instance should have been registred incrementally");
        }

        /// <summary>
        /// Even if Word won't display the 10th levels, the conversion should not fail
        /// </summary>
        [Test(Description = "Word doesn't display more than 8 deep levels.")]
        public void MaxNumberingLevel_ShouldBeIgnored()
        {
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i <= maxLevel; i++)
                sb.AppendFormat("<ol><li>Item {0}", i+1);
            for (int i = 0; i <= maxLevel; i++)
                sb.Append("</li></ol>");

            var elements = converter.Parse(sb.ToString());

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            var inst = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => i.AbstractNumId?.Val == absNum.AbstractNumberId)
                .SingleOrDefault();
            Assert.That(inst, Is.Not.Null);
            Assert.That(inst.NumberID?.Value, Is.Not.Null);

            Assert.That(elements, Has.Count.EqualTo(maxLevel + 1));
            Assert.That(elements.Cast<Paragraph>().Select(e => 
                e.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value),
                Has.All.EqualTo(inst.NumberID.Value),
                "All paragraphs are linked to the same list instance");
            Assert.That(elements.Last().GetFirstChild<ParagraphProperties>()?
                .NumberingProperties?.NumberingLevelReference?.Val?.Value, Is.EqualTo(maxLevel),
                "Level must be maxed out");
        }

        [Test(Description = "Apply Word document style on list scope")]
        public void UseVariantStyle_List_ReturnsAppliedStyle()
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithCustomStyle.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;
            HtmlConverter converter = new(mainPart);

            var elements = converter.Parse(@"<ul class='no-mapping-cls CustomStyle1'>
                <li>Item 1</li>
                <li>Item 2</li>
                <li>Item 3</li>
            </ul>");

            Assert.That(elements, Has.Count.EqualTo(3));
            Assert.That(elements.Cast<Paragraph>().Select(e => 
                e.ParagraphProperties?.ParagraphStyleId?.Val?.Value),
                Has.All.EqualTo("CustomStyle1"),
                "All paragraphs are linked to the same list instance");
        }

        [Test(Description = "Apply Word document style on list item scope")]
        public void UseVariantStyle_ListItem_ReturnsAppliedStyle()
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithCustomStyle.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;
            HtmlConverter converter = new(mainPart);

            var elements = converter.Parse(@"<ul>
                <li>Item 1</li>
                <li class='no-mapping-cls CustomStyle1'>Item 2</li>
                <li>Item 3</li>
            </ul>");

            var numbering = mainPart.NumberingDefinitionsPart?.Numbering;
            Assert.That(numbering, Is.Not.Null);
            Assert.That(elements, Has.Count.EqualTo(3));
            Assert.Multiple(() =>
            {
                Assert.That(elements.First().GetFirstChild<ParagraphProperties>()?.ParagraphStyleId?.Val?.Value,
                    Is.EqualTo(converter.HtmlStyles.DefaultStyles.ListParagraphStyle));
                Assert.That(elements.ElementAt(1).GetFirstChild<ParagraphProperties>()?.ParagraphStyleId?.Val?.Value,
                    Is.EqualTo("CustomStyle1"));
                Assert.That(elements.Last().GetFirstChild<ParagraphProperties>()?.ParagraphStyleId?.Val?.Value,
                    Is.EqualTo(converter.HtmlStyles.DefaultStyles.ListParagraphStyle));
            });
        }

        [Test(Description = "Resume indenting from existing numbering (default behaviour)")]
        public async Task ContinueNumbering_ReturnsSecondList_ContinueOrder()
        {
            await converter.ParseBody(@"<ol><li>Item 1</li></ol>");

            await converter.ParseBody("<ol><li>Item 2</li></ol>");

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            var instances = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => i.AbstractNumId!.Val == absNum.AbstractNumberId);
            Assert.That(instances, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(instances.Count(), Is.EqualTo(1));
                Assert.That(instances.Select(i => i.NumberID?.HasValue), Has.All.True);
            });

            var paragraphs = mainPart.Document.Body!.Elements<Paragraph>();
            Assert.That(paragraphs, Is.Not.Empty);
            Assert.That(paragraphs.Select(e => 
                e.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value),
                Has.All.EqualTo(instances.First().NumberID!.Value),
                "All paragraphs are linked to the same list instance");
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test(Description = "Stop indenting from existing numbering (issue #57)")]
        public async Task DisableContinueNumbering_ReturnsSecondList_RestartingOrder()
        {
            await converter.ParseBody(@"<ol><li>Item 1</li></ol>");

            converter.ContinueNumbering = false;
            await converter.ParseBody("<ol><li>Item 2</li></ol>");

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            var instances = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => i.AbstractNumId!.Val == absNum.AbstractNumberId);
            Assert.That(instances, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(instances.Count(), Is.EqualTo(2), "Expecting 2 distinct instances of the list");
                Assert.That(instances.Select(i => i.NumberID?.HasValue), Has.All.True);
            });

            var paragraphs = mainPart.Document.Body!.Elements<Paragraph>();
            Assert.That(paragraphs, Is.Not.Empty);
            Assert.That(paragraphs.Select(e => 
                e.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value),
                Is.Unique,
                "All paragraphs use different list instances");
            AssertThatOpenXmlDocumentIsValid();
        }

        /// <summary>
        /// Tiered numbering such as: 1, 1.1, 1.1.1
        /// </summary>
        [Test(Description = "Nested numbering (issue #81)")]
        public async Task DecimalTieredStyle_ReturnsListWithTieredNumbering()
        {
            await converter.ParseBody(
                @"<ol style='list-style-type:decimal-tiered'>
                    <li>Item 1
                        <ol><li>Item 1.1</li></ol>
                    </li>
                    <li>Item 2</li>
                </ol>");

            var elements = mainPart.Document.Body!.ChildElements;
            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            var instances = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => i.AbstractNumId!.Val == absNum.AbstractNumberId);
            Assert.That(instances, Is.Not.Null);
            var levels = absNum.Elements<Level>();
            Assert.Multiple(() =>
            {
                Assert.That(instances.Count(), Is.EqualTo(1));
                Assert.That(instances.Select(i => i.NumberID?.HasValue), Has.All.True);
                Assert.That(levels.Count(), Is.EqualTo(maxLevel + 1));
                Assert.That(levels.Select(l => l.NumberingFormat?.Val?.Value), Has.All.EqualTo(NumberFormatValues.Decimal));
                Assert.That(levels.Select(l => l.PreviousParagraphProperties?.Indentation?.Left?.Value), Has.All.EqualTo("0"),
                    "Decimal Tiered style must all be aligned on left with no indent");
            });

            Assert.That(elements, Is.Not.Empty);
            // exception rule: this style should cascade to nested lists
            Assert.That(elements.Cast<Paragraph>().Select(e => 
                e.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value),
                Has.All.EqualTo(instances.First().NumberID!.Value),
                "All paragraphs are linked to the same list instance");
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test(Description = "Allow to specify another start value for the first item of a `ol` list")]
        public void OverrideStartNumber_WithOl_ShouldSucceed()
        {
            const short startLevel = 3;
            var elements = converter.Parse($"<ol start='{startLevel}'><li>Item</li></ol>");
            Assert.That(elements, Is.Not.Empty);

            var inst = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>()
                .SingleOrDefault();
            Assert.That(inst, Is.Not.Null);
            Assert.That(inst.GetFirstChild<LevelOverride>()?.StartOverrideNumberingValue?.Val?.Value,
                Is.EqualTo(startLevel));
        }

        [Test(Description = "Allow to specify another start value for the first item of a `ul` list")]
        public void OverrideStartNumber_WithUl_ShouldBeIgnored()
        {
            var elements = converter.Parse($"<ul start='3'><li>Item</li></ul>");
            Assert.That(elements, Is.Not.Empty);

            var inst = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>()
                .SingleOrDefault();
            Assert.That(inst, Is.Not.Null);
            Assert.That(inst.GetFirstChild<LevelOverride>()?.StartOverrideNumberingValue?.Val?.Value,
                Is.EqualTo(1),
                "Start value should be ignored for `ul`");
        }

        [Test]
        public async Task RomanList_ReturnsListWithCustomStyle()
        {
            await converter.ParseBody(@"<ul style='list-style-type:lower-roman'>
                    <li>Item 1</li>
                </ul>");

            var elements = mainPart.Document.Body!.ChildElements;
            Assert.That(elements, Is.Not.Empty);
            Assert.That(elements, Is.All.TypeOf<Paragraph>());
            var numId = ((Paragraph) elements[0]).ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            Assert.That(numId, Is.Not.Null);

            var numInst = mainPart.NumberingDefinitionsPart!.Numbering
                .Elements<NumberingInstance>()
                .Single(i => i.NumberID?.Value == numId);
            Assert.That(numInst.AbstractNumId?.Val?.Value, Is.Not.Null);

            var absNums = mainPart.NumberingDefinitionsPart.Numbering
                .Elements<AbstractNum>();
            var absNum = absNums.FirstOrDefault(a => a.AbstractNumberId == numInst.AbstractNumId.Val);
            Assert.That(absNum, Is.Not.Null);
            Assert.That(absNum.NumberingStyleLink?.Val?.Value, Is.EqualTo("Harvard"));

            var style = mainPart.StyleDefinitionsPart!.Styles!
                .Elements<Style>()
                .FirstOrDefault(s => s.StyleName?.Val == "Harvard");
            Assert.That(style, Is.Not.Null);
            AssertThatOpenXmlDocumentIsValid();
        }

        [Test(Description = "Restart indenting in second nested numbering (issue #83)")]
        public async Task NestedNumbering_ReturnsNestedList_RestartingOrder()
        {
            await converter.ParseBody(@"
                <ol style='list-style-type: decimal;'>
                <li>item 1a
                    <ol style='list-style-type: lower-alpha;'>
                        <li>item 1.1a</li>
                    </ol>
                </li>
                </ol>
                <ol style='list-style-type: decimal;'>
                <li>item 1b
                    <ol style='list-style-type: lower-alpha;'>
                        <li>item 1.1b</li>
                    </ol>
                </li>
                </ol>");

            var elements = mainPart.Document.Body!.ChildElements;
            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>();
            Assert.That(absNum, Is.Not.Null);
            Assert.That(absNum.Count(), Is.EqualTo(2));

            var absNumIds = new HashSet<int>(absNum.Select(a => a.AbstractNumberId!.Value));
            var instances = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => absNumIds.Contains(i.AbstractNumId!.Val!));
            Assert.That(instances, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(instances.Count(), Is.EqualTo(4), "Expecting 4 distinct instances of the list");
                Assert.That(instances.Select(i => i.NumberID?.Value), Is.Unique);
            });
            Assert.That(instances.Last().GetFirstChild<LevelOverride>()?.LevelIndex?.Value, Is.EqualTo(1));
            Assert.That(instances.Last().GetFirstChild<LevelOverride>()?.StartOverrideNumberingValue?.Val?.Value, Is.EqualTo(1));
            AssertThatOpenXmlDocumentIsValid();
        }

        [TestCase("rtl", true)]
        [TestCase("ltr", false)]
        [TestCase("", null)]
        public void WithRtl_ReturnsBidi(string dir, bool? expectedValue)
        {
            var elements = converter.Parse($@"<ol dir='{dir}'>
                <li>Item 1</li><li>Item 2</li>
            </ol>");

            Assert.Multiple(() => {
                Assert.That(elements, Has.Count.EqualTo(2));
                Assert.That(elements, Is.All.TypeOf<Paragraph>());
                Assert.That(mainPart.NumberingDefinitionsPart?.Numbering, Is.Not.Null);
            });
            var bidis = elements.Cast<Paragraph>().Select(p => p.ParagraphProperties?.BiDi?.Val?.Value);
            Assert.That(bidis, Is.All.EqualTo(expectedValue));
        }

        [TestCase("rtl", "rtl", ExpectedResult = true)]
        [TestCase("rtl", "ltr", ExpectedResult = false)]
        [TestCase("rtl", "", ExpectedResult = true)]
        [TestCase("", "rtl", ExpectedResult = true)]
        public bool? WithNestedRtl_ReturnsBidi(string dir, string nestedDir)
        {
            var elements = converter.Parse($@"<ol dir='{dir}'>
                    <li>Item 1
                        <ol dir='{nestedDir}'><li>Item 1.1</li></ol>
                    </li>
                </ol>");

            Assert.Multiple(() => {
                Assert.That(elements, Has.Count.EqualTo(2));
                Assert.That(elements, Is.All.TypeOf<Paragraph>());
                Assert.That(mainPart.NumberingDefinitionsPart?.Numbering, Is.Not.Null);
            });
            var bidi = elements.Last().GetFirstChild<ParagraphProperties>()?.BiDi;
            return bidi?.Val?.Value;
        }

        [Test]
        public void NestedNumberList_ReturnsIncrementalIndentation()
        {
            const int maxLevel = 8;
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < maxLevel; i++)
                sb.AppendFormat("<ol><li>Item {0}", i+1);
            for (int i = 0; i < maxLevel; i++)
                sb.Append("</li></ol>");

            converter.Parse(sb.ToString());

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            var levels = absNum.Elements<Level>();
            Assert.That(levels.Count(), Is.EqualTo(maxLevel + 1));
            for (int i = 0; i <= maxLevel; i++)
            {
                var level = levels.ElementAt(i);
                var ident = level.PreviousParagraphProperties?.Indentation;
                Assert.That(ident?.Hanging?.Value, Is.EqualTo("360"));
                Assert.That(Convert.ToInt32(ident?.Left?.Value), Is.EqualTo((i + 1) * 2 * 360));
                TestContext.Out.WriteLine($"{i}. {ident?.Left?.Value}");
            }
        }

        [Test(Description = "Nested list must be a children of a `li` tag but some editor are not respecting the W3C standard (issue #173)")]
        public async Task NestedNumberList_NonCompliant_ReturnsIncrementalIndentation()
        {
            await converter.ParseBody(@"<ol>
                <li>Item1</li>
                <li>Item2</li>
                    <ol><li>Item 2.1</li></ol>
                </ol>");

            var absNum = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<AbstractNum>()
                .SingleOrDefault();
            Assert.That(absNum, Is.Not.Null);

            var inst = mainPart.NumberingDefinitionsPart?.Numbering
                .Elements<NumberingInstance>().Where(i => i.AbstractNumId?.Val == absNum.AbstractNumberId)
                .SingleOrDefault();
            Assert.That(inst, Is.Not.Null);
            Assert.That(inst.NumberID?.Value, Is.Not.Null);

            var elements = mainPart.Document.Body!.ChildElements;
            Assert.Multiple(() => {
                Assert.That(elements, Has.Count.EqualTo(3));
                Assert.That(elements, Is.All.TypeOf<Paragraph>());
                Assert.That(mainPart.NumberingDefinitionsPart?.Numbering, Is.Not.Null);
            });

            // assert paragraphs linked to numbering instance
            Assert.Multiple(() =>
            {
                Assert.That(elements.Cast<Paragraph>().Select(e => 
                    e.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value),
                    Has.All.EqualTo(inst.NumberID.Value),
                    "All paragraphs are linked to the same list instance");
                Assert.That(elements.Take(2).Select(p => p.GetFirstChild<ParagraphProperties>()?.NumberingProperties?.NumberingLevelReference?.Val?.Value), Has.All.EqualTo(0));
                Assert.That(elements.Last().GetFirstChild<ParagraphProperties>()?.NumberingProperties?.NumberingLevelReference?.Val?.Value, Is.EqualTo(1));
            });
            AssertThatOpenXmlDocumentIsValid();
        }
    }
}