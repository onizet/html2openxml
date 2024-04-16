/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the ordered <c>ol</c> and unordered <c>ul</c> list.
/// </summary>
sealed class ListExpression(IHtmlElement node) : FlowElementExpression(node)
{
#if NET5_0_OR_GREATER
    readonly record struct ListContext(string Name, int AbsNumId, int InstanceId, int Level);
#else
    readonly struct ListContext(string listName, int absNumId, int instanceId, int level)
    {
        public readonly string Name = listName;
        public readonly int AbsNumId = absNumId;
        public readonly int InstanceId = instanceId;
        public readonly int Level = level;
    }
#endif

    /// <summary>Hard-coded value from Word</summary>
    const int MAX_LEVEL = 8;
    const string HEADING_NUMBERING_NAME = "decimal-heading-multi";
    // https://www.w3schools.com/cssref/playdemo.php?filename=playcss_list-style-type
    // https://answers.microsoft.com/en-us/msoffice/forum/all/custom-list-number-style/21a54399-4404-4c37-8843-2ccaaf827485
    // Image bullet: http://officeopenxml.com/WPnumbering-imagesAsSymbol.php
    private static readonly HashSet<string> supportedListTypes = 
        ["disc", "decimal", "square", "circle",
         "lower-alpha", "upper-alpha", "lower-latin", "upper-latin",
         "lower-roman", "upper-roman"];
    private static readonly IDictionary<string, AbstractNum> predefinedNumberingLists = InitKnownLists();
    /// <summary>Contains the list of templated list along with the AbstractNumbId</summary>
    private Dictionary<string, int>? knownAbsNumIds;
    /// <summary>Contains the list of numbering instance.</summary>
    private Dictionary<int, int>? knownInstanceIds;
    private Numbering? numbering;


    public override IEnumerable<OpenXmlCompositeElement> Interpret(ParsingContext context)
    {
        var liNodes = node.Children.Where(n => n.LocalName == "li");
        if (!liNodes.Any()) yield break;

        // Ensure the numbering.xml file exists or any numbering or bullets list will results
        // in simple numbering list (1.   2.   3...)
        NumberingDefinitionsPart numberingPart = context.MainPart.NumberingDefinitionsPart
            ?? context.MainPart.AddNewPart<NumberingDefinitionsPart>();

        if (numberingPart.Numbering == null)
        {
            new Numbering().Save(numberingPart);
        }

        numbering = context.MainPart.NumberingDefinitionsPart!.Numbering;
        ListContext? parentContext = null;
        var listContext = context.Properties<ListContext>("listContext");
        var listStyle = GetListType(node);
        if (listContext.InstanceId == 0 || listContext.Name != listStyle)
        {
            InitNumberingIds(context);

            parentContext = listContext;
            var abstractNumId = FindListTemplate(context, listStyle);
            listContext = ConcretiseInstance(context, abstractNumId, listStyle, listContext.Level);

            numbering.Append(
                new NumberingInstance(
                    new AbstractNumId() { Val = listContext.AbsNumId },
                    new LevelOverride(
                        new StartOverrideNumberingValue() { Val = 1 }
                    )
                )
                { NumberID = listContext.InstanceId });
        }
        else
        {
            parentContext = listContext;
            listContext = new ListContext(listContext.Name, listContext.AbsNumId, 
                listContext.InstanceId, listContext.Level + 1);
        }

        context.Properties("listContext", listContext);

        // +1 because index starts on 1 and not 0
        var level = Math.Min(listContext.Level, MAX_LEVEL+1);
        foreach (IHtmlElement liNode in liNodes.Cast<IHtmlElement>())
        {
            var expression = new FlowElementExpression(liNode);
            var childElements = expression.Interpret(context);
            Paragraph p = (Paragraph) childElements.First();

            p.InsertInProperties(prop => {
                //todo: GetStyleIdForListItem
                prop.ParagraphStyleId = context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.ListParagraphStyle);
                prop.Indentation = level < 2? null : new() { Left = (level * 720).ToString(CultureInfo.InvariantCulture) };
                prop.NumberingProperties = new NumberingProperties {
                    NumberingLevelReference = new() { Val = level - 1 },
                    NumberingId = new() { Val = listContext.InstanceId }
                };
            });

            foreach (var child in childElements)
                yield return child;
        }

        if (parentContext.HasValue)
        {
             context.Properties("listContext", parentContext.Value);
        }
        else
        {
            context.Properties("listContext", null);
        }
    }

    /// <summary>
    /// Resolve the list style to determine which NumberList style to apply.
    /// </summary>
    private static string GetListType(IElement listNode)
    {
        var styleAttributes = HtmlAttributeCollection.ParseStyle(listNode.GetAttribute("style"));
        string? type = styleAttributes["list-style-type"];

        if (string.IsNullOrEmpty(type) || !supportedListTypes.Contains(type!))
        {
            bool orderedList = listNode.NodeName.Equals(TagNames.Ol, StringComparison.OrdinalIgnoreCase);
            type = orderedList? "decimal" : "disc";
        }

        return type!;
    }

    /// <summary>
    /// Resolve the next available <see cref="AbstractNum.AbstractNumberId"/> (they must be unique and ordered).
    /// </summary>
    private static int IncrementAbstractNumId(ParsingContext context, Numbering numbering)
    {
        var absNumIdRef = context.Properties<int?>("absNumIdRef");
        if (!absNumIdRef.HasValue)
        {
            absNumIdRef = 0;
            // The absNumIdRef Id is a required field and should be unique. We will loop through the existing Numbering definition
            // to retrieve the highest Id and reconstruct our own list definition template.
            foreach (var abs in numbering.Elements<AbstractNum>())
            {
                if (abs.AbstractNumberId != null && abs.AbstractNumberId > absNumIdRef)
                    absNumIdRef = abs.AbstractNumberId;
            }
        }

        // compute the next list instance ID seed. We start at 1 because 0 has a special meaning: 
        // The w:numId can contain a value of 0, which is a special value that indicates that numbering was removed
        // at this level of the style hierarchy. While processing this markup, if the w:val='0',
        // the paragraph does not have a list item (http://msdn.microsoft.com/en-us/library/ee922775(office.14).aspx)
        absNumIdRef++;
        context.Properties("absNumIdRef", absNumIdRef);
        return absNumIdRef.Value;
    }

    /// <summary>
    /// Resolve the next available <see cref="NumberingInstance.NumberID"/> (they must be unique and ordered).
    /// </summary>
    private static int IncrementInstanceId(ParsingContext context, Numbering numbering)
    {
        var instanceId = context.Properties<int?>("listInstanceId");
        if (!instanceId.HasValue)
        {
            instanceId = 1;

            // compute the next list instance ID seed. We start at 1 because 0 has a special meaning: 
            // The w:numId can contain a value of 0, which is a special value that indicates that numbering was removed
            // at this level of the style hierarchy. While processing this markup, if the w:val='0',
            // the paragraph does not have a list item (http://msdn.microsoft.com/en-us/library/ee922775(office.14).aspx)
            foreach (NumberingInstance inst in numbering.Elements<NumberingInstance>())
            {
                if (inst.NumberID?.Value > instanceId) instanceId = inst.NumberID;
            }
        }
        instanceId++;
        context.Properties("listInstanceId", instanceId);
        return instanceId.Value;
    }

    /// <summary>
    /// Discover the list of existing templates and instances, already registred in the document.
    /// </summary>
    private void InitNumberingIds(ParsingContext context)
    {
        knownAbsNumIds = context.Properties<Dictionary<string, int>>("knownAbsNumIds");
        knownInstanceIds = context.Properties<Dictionary<int, int>>("knownInstanceIds");
        if (knownAbsNumIds != null && knownInstanceIds != null) return;

        knownAbsNumIds = [];
        knownInstanceIds = [];
        int absNumIdRef = 0;

        // The absNumIdRef Id is a required field and should be unique. We will loop through the existing Numbering definition
        // to retrieve the highest Id and reconstruct our own list definition template.
        foreach (var abs in numbering!.Elements<AbstractNum>())
        {
            if (abs.AbstractNumberId != null && abs.AbstractNumberId > absNumIdRef)
                absNumIdRef = abs.AbstractNumberId;
        }
        absNumIdRef++;

        IEnumerable<AbstractNum> existingAbstractNums = numbering.ChildElements
            .Where(e => e != null && e is AbstractNum).Cast<AbstractNum>();

        knownAbsNumIds = existingAbstractNums
            .Where(a => a.AbstractNumDefinitionName != null && a.AbstractNumDefinitionName.Val != null)
            .ToDictionary(a => a.AbstractNumDefinitionName!.Val!.Value!, a => a.AbstractNumberId!.Value);

        foreach (NumberingInstance inst in numbering.Elements<NumberingInstance>())
        {
            knownInstanceIds.Add(inst.AbstractNumId!.Val!.Value, inst.NumberID!.Value);
        }

        context.Properties("knownAbsNumIds", knownAbsNumIds);
        context.Properties("knownInstanceIds", knownInstanceIds);
    }

    /// <summary>
    /// Predefined template of lists.
    /// </summary>
    private static Dictionary<string, AbstractNum> InitKnownLists()
    {
        var knownAbstractNums = new Dictionary<string, AbstractNum>();

        // This minimal numbering definition has been inspired by the documentation OfficeXMLMarkupExplained_en.docx
        // http://www.microsoft.com/downloads/details.aspx?FamilyID=6f264d0b-23e8-43fe-9f82-9ab627e5eaa3&displaylang=en
        foreach (var (listName, formatValue, text) in new[] {
            ("decimal", NumberFormatValues.Decimal, "%{0}."),
            ("disc", NumberFormatValues.Bullet, "•"),
            ("square", NumberFormatValues.Bullet, "▪"),
            ("circle", NumberFormatValues.Bullet, "o"),
            ("upper-alpha", NumberFormatValues.UpperLetter, "%{0}."),
            ("lower-alpha", NumberFormatValues.LowerLetter, "%{0}."),
            ("upper-roman", NumberFormatValues.UpperRoman, "%{0}."),
            ("lower-roman", NumberFormatValues.LowerRoman, "%{0}."),
            ("upper-greek", NumberFormatValues.UpperLetter, "%{0}."),
            ("lower-greek", NumberFormatValues.LowerLetter, "%{0}."),
        })
        {
            var abstractNum = new AbstractNum(
                new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel },
                new RunProperties(
                    new RunFonts() { HighAnsi = "Arial Unicode MS" }
                )
            ) { AbstractNumDefinitionName = new() { Val = listName } };

            bool useSymbol = listName.EndsWith("-greek");
            for (var lvlIndex = 0; lvlIndex <= MAX_LEVEL; lvlIndex++)
            {
                abstractNum.Append(new Level {
                    StartNumberingValue = new() { Val = 1 },
                    NumberingFormat = new() { Val = formatValue },
                    LevelIndex = lvlIndex,
                    LevelText = new() { Val = string.Format(text, lvlIndex) },
                    LevelJustification = new() { Val = LevelJustificationValues.Left },
                    PreviousParagraphProperties = new() {
                        Indentation = new() { Left = "720", Hanging = "360" }
                    },
                    NumberingSymbolRunProperties = useSymbol? new () {
                        RunFonts = new() { Ascii = "Symbol", Hint = FontTypeHintValues.Default }
                    } : null
                });
            }

            knownAbstractNums.Add(listName, abstractNum);
        }

        // decimal-heading-multi
        // WARNING: only use this for headings
        knownAbstractNums.Add(HEADING_NUMBERING_NAME, new AbstractNum(
            new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
            new Level {
                StartNumberingValue = new StartNumberingValue() { Val = 1 },
                NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
                LevelIndex = 0,
                LevelText = new LevelText() { Val = "%1." }
            }
        ) { AbstractNumDefinitionName = new() { Val = HEADING_NUMBERING_NAME } });

        return knownAbstractNums;
    }

    /// <summary>
    /// Find or register an list template from the document.
    /// </summary>
    private int FindListTemplate(ParsingContext context, string listName, bool cascading = false)
    {
        // lookup for a predefined list style in the template collection
        if (knownAbsNumIds!.TryGetValue(listName, out int abstractNumId))
        {
            return abstractNumId;
        }

        Numbering numberingPart = context.MainPart.NumberingDefinitionsPart!.Numbering;

        // at this stage, we have sanitized the list style so it's safe to grab them from the predefined template lists
        var abstractNum =  predefinedNumberingLists[listName];
        abstractNum = (AbstractNum) abstractNum.CloneNode(true);
        abstractNum.AbstractNumberId = IncrementAbstractNumId(context, numberingPart);
        var level1 = abstractNum.GetFirstChild<Level>()!;

        /*Level level1 = abstractNum.GetFirstChild<Level>()!;
        // skip the first level, starts to 2
        for (int i = 1; i < maxDepth.Depth; i++)
        {
            Level level = new() {
                StartNumberingValue = new StartNumberingValue() { Val = 1 },
                NumberingFormat = new NumberingFormat() { Val = level1.NumberingFormat?.Val },
                LevelIndex = i - 1
            };

            if (cascading) 
            {
                // if we're cascading, that means we don't want any indentation 
                // + our leveltext should contain the previous levels as well
                var lvlText = new System.Text.StringBuilder();

                for (int lvlIndex = 1; lvlIndex <= i; lvlIndex++)
                    lvlText.AppendFormat("%{0}.", lvlIndex);

                level.LevelText = new LevelText() { Val = lvlText.ToString() };
            }
            else
            {
                level.LevelText = new LevelText() { Val = $"%{i}." };
                level.PreviousParagraphProperties = 
                    new PreviousParagraphProperties {
                        Indentation = new() { Left = (720 * i).ToString(CultureInfo.InvariantCulture), Hanging = "360" }
                    };
            }

            abstractNum.AppendChild(level);
        }*/

        // this is not documented but MS Word needs that all the AbstractNum are stored consecutively.
        // Otherwise, it will apply the "NoList" style to the existing ListInstances.
        // This is the reason why I insert all the items after the last AbstractNum.
        var lastAbsNum = numberingPart.GetLastChild<AbstractNum>();
        if (lastAbsNum == null)
            numberingPart.InsertAt(abstractNum, 0);
        else
            numberingPart.InsertAfter(abstractNum, lastAbsNum);

        abstractNumId = abstractNum.AbstractNumberId;

        // For Roman numbering (I, II, i, ii), we must define a dedicated style
        // and link it to the numbering definition
        if (level1.NumberingFormat!.Val! == NumberFormatValues.LowerRoman
            || level1.NumberingFormat.Val! == NumberFormatValues.UpperRoman)
        {
            abstractNumId++;
            numberingPart.InsertAfter(new AbstractNum{
                AbstractNumberId = abstractNumId,
                MultiLevelType = new MultiLevelType { Val = MultiLevelValues.HybridMultilevel },
                NumberingStyleLink =  new NumberingStyleLink { Val = "Harvard" }
            }, lastAbsNum);

            abstractNum.StyleLink = new StyleLink { Val = "Harvard" };
            context.DocumentStyle.AddStyle("Harvard", new Style (
                new Name { Val = "Harvard" },
                new ParagraphProperties(
                    new NumberingProperties() { NumberingId = new() { Val = abstractNum.AbstractNumberId } }
                )) {
                Type = StyleValues.Numbering,
                StyleId = "Harvard"
            });
        }

        // register this new template
        knownAbsNumIds.Add(listName, abstractNumId);
        context.Properties("knownAbsNumIds", knownAbsNumIds);

        return abstractNumId;
    }

    private ListContext ConcretiseInstance(ParsingContext context, int abstractNumId, string listStyle, int currentLevel)
    {
        if (!knownInstanceIds!.TryGetValue(abstractNumId, out int instanceId))
        {
            // create a new instance of that list template
            instanceId = IncrementInstanceId(context, numbering!);
            knownInstanceIds.Add(abstractNumId, instanceId);
        }
        else
            // if the previous element is the same list style,
            // we must restart the ordering to 0
            if (node.PreviousElementSibling != null &&
            (node.PreviousElementSibling.LocalName == "ol" ||
             node.PreviousElementSibling.LocalName == "ul")
             && GetListType(node.PreviousElementSibling) == listStyle)
        {
            instanceId = IncrementInstanceId(context, numbering!);
            return new ListContext(listStyle, abstractNumId, instanceId, 1);
        }

        return new ListContext(listStyle, abstractNumId, instanceId, currentLevel + 1);
    }
}