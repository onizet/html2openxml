using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;

namespace NotesFor.HtmlToOpenXml
{
	sealed class NumberingListStyleCollection
	{
        private MainDocumentPart mainPart;
        private int currentNumId, absNumId, numberLevelRef;
        private bool firstItem;
        private string levelText;
        private NumberFormatValues format;


        public NumberingListStyleCollection(MainDocumentPart mainPart)
        {
            this.mainPart = mainPart;
            InitNumberingIds();
        }


        #region InitNumberingIds

        private void InitNumberingIds()
        {
            // Ensure the numbering.xml file exists or any numbering or bullets list will results
            // in simple numbering list (1.   2.   3...)
            if (mainPart.NumberingDefinitionsPart == null || mainPart.NumberingDefinitionsPart.Numbering == null)
            {
                // This minimal numbering definition has been inspired by the documentation OfficeXMLMarkupExplained_en.docx
                // http://www.microsoft.com/downloads/details.aspx?FamilyID=6f264d0b-23e8-43fe-9f82-9ab627e5eaa3&displaylang=en

                NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                new Numbering(
                    //8 kinds of abstractnum.
                     new AbstractNum(
                        new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new NumberingFormat() { Val = NumberFormatValues.Decimal },
                            new LevelText() { Val = "%1." },
                            new PreviousParagraphProperties(
                                new Indentation() { Left = "420", Hanging = "360" })
                        ) { LevelIndex = 0 }
                    ) { AbstractNumberId = 0 },
                    new AbstractNum(
                        new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                        new Level(
                            new NumberingFormat() { Val = NumberFormatValues.Bullet },
                            new LevelText() { Val = "•" },
                            new PreviousParagraphProperties(
                                new Indentation() { Left = "420", Hanging = "360" })
                        ) { LevelIndex = 0 }
                    ) { AbstractNumberId = 1 },
                    new AbstractNum(
                        new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                        new Level(
                            new NumberingFormat() { Val = NumberFormatValues.Bullet },
                            new LevelText() { Val = "■" },
                            new PreviousParagraphProperties(
                                new Indentation() { Left = "420", Hanging = "360" })
                        ) { LevelIndex = 0 }
                    ) { AbstractNumberId = 2 },
                    new AbstractNum(
                        new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                        new Level(
                            new NumberingFormat() { Val = NumberFormatValues.Bullet },
                            new LevelText() { Val = "o" },
                            new PreviousParagraphProperties(
                                new Indentation() { Left = "420", Hanging = "360" })
                        ) { LevelIndex = 0 }
                    ) { AbstractNumberId = 3 },
                    new AbstractNum(
                        new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new NumberingFormat() { Val = NumberFormatValues.UpperLetter },
                            new LevelText() { Val = "%1." },
                            new PreviousParagraphProperties(
                                new Indentation() { Left = "420", Hanging = "360" })
                        ) { LevelIndex = 0 }
                    ) { AbstractNumberId = 4 },
                    new AbstractNum(
                        new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new NumberingFormat() { Val = NumberFormatValues.LowerLetter },
                            new LevelText() { Val = "%1." },
                            new PreviousParagraphProperties(
                                new Indentation() { Left = "420", Hanging = "360" })
                        ) { LevelIndex = 0 }
                    ) { AbstractNumberId = 5 },
                    new AbstractNum(
                        new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new NumberingFormat() { Val = NumberFormatValues.UpperRoman },
                            new LevelText() { Val = "%1." },
                            new PreviousParagraphProperties(
                                new Indentation() { Left = "420", Hanging = "360" })
                        ) { LevelIndex = 0 }
                    ) { AbstractNumberId = 6 },
                    new AbstractNum(
                        new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new NumberingFormat() { Val = NumberFormatValues.LowerRoman },
                            new LevelText() { Val = "%1." },
                            new PreviousParagraphProperties(
                                new Indentation() { Left = "420", Hanging = "360" })
                        ) { LevelIndex = 0 }
                    ) { AbstractNumberId = 7 }
                    ).Save(numberingPart);
            }
        }

        #endregion

        #region BeginList

        public void BeginList(HtmlEnumerator en)
        {
            if (en.CurrentTag.Equals("<ul>", StringComparison.InvariantCultureIgnoreCase))
            {
                switch (en.StyleAttributes["list-style-type"])
                {
                    case "none":
                    case "circle":
                        format = NumberFormatValues.Bullet;
                        levelText = "o";
                        absNumId = 3;
                        break;
                    case "square":
                        format = NumberFormatValues.Bullet;
                        levelText = "■";
                        absNumId = 2;
                        break;
                    case "disc":
                    default:
                        format = NumberFormatValues.Bullet;
                        levelText = "•";
                        absNumId = 1;
                        break;
                }
            }
            else
            {
                switch (en.StyleAttributes["list-style-type"])
                {
                    case "upper-alpha":
                        format = NumberFormatValues.UpperLetter;
                        levelText = "%1.";
                        absNumId = 4;
                        break;
                    case "lower-alpha":
                        format = NumberFormatValues.LowerLetter;
                        levelText = "%1.";
                        absNumId = 5;
                        break;
                    case "upper-roman":
                        format = NumberFormatValues.UpperRoman;
                        levelText = "%1.";
                        absNumId = 6;
                        break;
                    case "lower-roman":
                        format = NumberFormatValues.LowerRoman;
                        levelText = "%1.";
                        absNumId = 7;
                        break;
                    case "decimal-leading-zero":
                    default:
                        format = NumberFormatValues.Decimal;
                        levelText = "%1.";
                        absNumId = 0;
                        break;
                }
            }

            numberLevelRef++;

            Numbering numbering = mainPart.NumberingDefinitionsPart.Numbering;
            numbering.Append(
                new NumberingInstance(
                    new AbstractNumId() { Val = absNumId }
                ) { NumberID = currentNumId });
            numbering.Save(mainPart.NumberingDefinitionsPart);
            numbering.Reload();
        }

        #endregion

        #region EndList

        public void EndList()
        {
            numberLevelRef--;
            firstItem = true;
        }

        #endregion

        public int ProcessItem(HtmlEnumerator en)
        {
            if (!firstItem) return currentNumId;

            firstItem = false;
            Int32 leftMarginSize = 0;
            if (en.StyleAttributes["margin-left"] != null)
            {
                Unit margin = en.StyleAttributes.GetAsUnit("margin-left");
                if (margin.IsValid)
                {
                    if (margin.Value > 0 && margin.Type == "px")
                        leftMarginSize = margin.Value;
                }
            }
            else if (en.StyleAttributes["margin"] != null)
            {
                Margin margin = en.StyleAttributes.GetAsMargin("margin");
                if (margin.IsValid && margin.Left.Value > 0 && margin.Left.Type == "px")
                {
                    leftMarginSize = margin.Left.Value;
                }
            }

            if (leftMarginSize > 0)
            {
                currentNumId++;
                Margin margin = en.StyleAttributes.GetAsMargin("margin");

                mainPart.NumberingDefinitionsPart.Numbering.Append(
                    new AbstractNum(
                            new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                            new Level(
                                new StartNumberingValue() { Val = 1 },
                                new NumberingFormat() { Val = format },
                                new LevelText() { Val = levelText },
                                new PreviousParagraphProperties(
                                    new Indentation() { Left = leftMarginSize.ToString(CultureInfo.InvariantCulture), Hanging = "360" })
                            ) { LevelIndex = 0 }
                        ) { AbstractNumberId = currentNumId + 8 });
                mainPart.NumberingDefinitionsPart.Numbering.Save(mainPart.NumberingDefinitionsPart);
                mainPart.NumberingDefinitionsPart.Numbering.Append(
                    new NumberingInstance(
                            new AbstractNumId() { Val = currentNumId + 8 }
                        ) { NumberID = currentNumId });
                mainPart.NumberingDefinitionsPart.Numbering.Save(mainPart.NumberingDefinitionsPart);
                mainPart.NumberingDefinitionsPart.Numbering.Reload();
            }

            return currentNumId;
        }

        //____________________________________________________________________
        //
        // Properties Implementation

        public Int32 LevelRef
        {
            get { return numberLevelRef; }
        }
    }
}