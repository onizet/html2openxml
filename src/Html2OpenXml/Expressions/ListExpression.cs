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
    readonly record struct ListDepth(int Depth, string Topology);
    readonly record struct ListContext(int AbsNumId, int InstanceId, int Level);
#else
    readonly struct ListDepth(int depth, string topology)
    {
        public readonly int Depth = depth;
        public readonly string Topology = topology;
    }
    readonly struct ListContext(int absNumId, int instanceId, int level)
    {
        public readonly int AbsNumId = absNumId;
        public readonly int InstanceId = instanceId;
        public readonly int Level = level;
    }
#endif

    const string HEADING_NUMBERING_NAME = "decimal-heading-multi";
    //TODO: extend to support greek, hebrew, hiragana, katakana, etc
    //https://www.w3schools.com/cssref/playdemo.php?filename=playcss_list-style-type
    //https://answers.microsoft.com/en-us/msoffice/forum/all/custom-list-number-style/21a54399-4404-4c37-8843-2ccaaf827485
    //Image bullet: http://officeopenxml.com/WPnumbering-imagesAsSymbol.php
    private static readonly HashSet<string> supportedListTypes = 
        ["disc", "decimal", "square", "circle", "decimal-leading-zero",
         "lower-alpha", "upper-alpha", "lower-latin", "upper-latin",
         "lower-roman", "upper-roman"];
    private static readonly IDictionary<string, AbstractNum> predefinedNumberingLists = InitKnownLists();
    /// <summary>Contains the list of templated list along with the AbstractNumbId</summary>
    private Dictionary<string, int>? knownAbsNumIds;


    public override IEnumerable<OpenXmlCompositeElement> Interpret(ParsingContext context)
    {
        var liNodes = node.QuerySelectorAll("li");
        if (!liNodes.Any()) yield break;

        Numbering numbering;

        var listContext = context.Properties<ListContext>("listContext");
        if (listContext.InstanceId == 0)
        {
            var nestedLists = EnumerateNestedLists(node, new ListDepth(1, GetListType(node)));
            var maxDepth = nestedLists.OrderByDescending(d => d.Depth).First();

            knownAbsNumIds = InitNumberingIds(context);
            // lookup for a predefined list style in the template collection
            if (!knownAbsNumIds.TryGetValue(maxDepth.Topology, out int numberingId))
            {
                numberingId = RegisterNewNumbering(context, maxDepth);
            }

            numbering = context.MainPart.NumberingDefinitionsPart!.Numbering;
            var instanceId = IncrementInstanceId(context, numbering);
            listContext = new ListContext(numberingId, instanceId, 1);

            numbering.Append(
                new NumberingInstance(
                    new AbstractNumId() { Val = listContext.AbsNumId },
                    new LevelOverride(
                        new StartOverrideNumberingValue() { Val = 1 }
                    )
                    //TOOD:{ LevelIndex = 0 }
                )
                { NumberID = listContext.InstanceId });
        }
        else
        {
            numbering = context.MainPart.NumberingDefinitionsPart!.Numbering;
            listContext = new ListContext(listContext.AbsNumId, listContext.InstanceId, listContext.Level + 1);
        }

        context.Properties("listContext", listContext);

        var level = listContext.Level;
        foreach (IHtmlElement liNode in liNodes.Cast<IHtmlElement>())
        {
            var expression = new FlowElementExpression(liNode);
            var childElements = expression.Interpret(context);
            Paragraph p = (Paragraph) childElements.First();

            p.InsertInProperties(prop => {
                //todo: GetStyleIdForListItem
                prop.ParagraphStyleId = context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.ListParagraphStyle);
                prop.Indentation = level < 2? null : new() { Left = (level * 780).ToString(CultureInfo.InvariantCulture) };
                prop.NumberingProperties = new NumberingProperties {
                    NumberingLevelReference = new NumberingLevelReference() { Val = level - 1 },
                    NumberingId = new NumberingId() { Val = listContext.AbsNumId }
                };
            });

            foreach (var child in childElements)
                yield return child;
        }

        if (level > 1)
        {
            listContext = new ListContext(listContext.AbsNumId, listContext.InstanceId, level - 1);
            context.Properties("listContext", listContext);
        }
    }

    /// <summary>
    /// Walk through the whole list hierarchy to provide a deep analysis of the structure.
    /// </summary>
    private static IEnumerable<ListDepth> EnumerateNestedLists(IElement node, ListDepth depth)
    {
        yield return depth;

        foreach (var nestedList in node.QuerySelectorAll("ul,ol"))
        {
            var childDepth = new ListDepth(depth.Depth + 1, depth.Topology + "+" + GetListType(nestedList));
            foreach (var list in EnumerateNestedLists(nestedList, childDepth))
                yield return list;
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

    private static Dictionary<string, int> InitNumberingIds(ParsingContext context)
    {
        var knownAbsNumIds = context.Properties<Dictionary<string, int>>("knownAbsNumIds");
        if (knownAbsNumIds != null) return knownAbsNumIds;

        knownAbsNumIds = [];
        int absNumIdRef = 0;

        // Ensure the numbering.xml file exists or any numbering or bullets list will results
        // in simple numbering list (1.   2.   3...)
        NumberingDefinitionsPart numberingPart = context.MainPart.NumberingDefinitionsPart
            ?? context.MainPart.AddNewPart<NumberingDefinitionsPart>();

        if (numberingPart.Numbering == null)
        {
            new Numbering().Save(numberingPart);
        }
        else
        {
            // The absNumIdRef Id is a required field and should be unique. We will loop through the existing Numbering definition
            // to retrieve the highest Id and reconstruct our own list definition template.
            foreach (var abs in numberingPart.Numbering.Elements<AbstractNum>())
            {
                if (abs.AbstractNumberId != null && abs.AbstractNumberId > absNumIdRef)
                    absNumIdRef = abs.AbstractNumberId;
            }
            absNumIdRef++;
        }

        // Check if we have already initialized our abstract nums
        // if that is the case, we should not add them again.
        // This supports a use-case where the HtmlConverter is called multiple times
        // on document generation, and needs to continue existing lists
        /*bool addNewAbstractNums = true;
        IEnumerable<AbstractNum> existingAbstractNums = numberingPart.Numbering!.ChildElements.Where(e => e != null && e is AbstractNum).Cast<AbstractNum>();

        if (existingAbstractNums.Count() >= absNumChildren.Length) // means we might have added our own already
        {
            foreach (var abstractNum in absNumChildren)
            {
                // Check if we can find this in the existing document
                addNewAbstractNums = addNewAbstractNums 
                    || !existingAbstractNums.Any(a => a.AbstractNumDefinitionName?.Val?.Value == abstractNum.AbstractNumDefinitionName?.Val?.Value);
            }
        }

        if (addNewAbstractNums)
        {
            // this is not documented but MS Word needs that all the AbstractNum are stored consecutively.
            // Otherwise, it will apply the "NoList" style to the existing ListInstances.
            // This is the reason why I insert all the items after the last AbstractNum.
            int lastAbsNumIndex = 0;
            if (absNumIdRef > 0)
            {
                lastAbsNumIndex = numberingPart.Numbering.ChildElements.Count-1;
                for (; lastAbsNumIndex >= 0; lastAbsNumIndex--)
                {
                    if(numberingPart.Numbering.ChildElements[lastAbsNumIndex] is AbstractNum)
                        break;
                }
            }

            lastAbsNumIndex = Math.Max(lastAbsNumIndex, 0);

            for (int i = 0; i < absNumChildren.Length; i++)
                numberingPart.Numbering.InsertAt(absNumChildren[i], i + lastAbsNumIndex);

            knownAbsNumIds = absNumChildren
                .ToDictionary(a => a.AbstractNumDefinitionName!.Val!.Value!, a => a.AbstractNumberId!.Value);
        } 
        else
        {*/
        IEnumerable<AbstractNum> existingAbstractNums = numberingPart.Numbering!.ChildElements
            .Where(e => e != null && e is AbstractNum).Cast<AbstractNum>();

        knownAbsNumIds = existingAbstractNums
            .Where(a => a.AbstractNumDefinitionName != null && a.AbstractNumDefinitionName.Val != null)
            .ToDictionary(a => a.AbstractNumDefinitionName!.Val!.Value!, a => a.AbstractNumberId!.Value);
        //}

        // compute the next list instance ID seed. We start at 1 because 0 has a special meaning: 
        // The w:numId can contain a value of 0, which is a special value that indicates that numbering was removed
        // at this level of the style hierarchy. While processing this markup, if the w:val='0',
        // the paragraph does not have a list item (http://msdn.microsoft.com/en-us/library/ee922775(office.14).aspx)
        //TODO:store this
        /*int nextInstanceID = 1;
        foreach (NumberingInstance inst in numberingPart.Numbering.Elements<NumberingInstance>())
        {
            if (inst.NumberID?.Value > nextInstanceID) nextInstanceID = inst.NumberID;
        }*/
        //numInstances.Push(new KeyValuePair<int, int>(nextInstanceID, -1));

        numberingPart.Numbering.Save();
        context.Properties("knownAbsNumIds", knownAbsNumIds);
        return knownAbsNumIds;
    }

    /// <summary>
    /// Predefined template of lists.
    /// </summary>
    private static IDictionary<string, AbstractNum> InitKnownLists()
    {
        var knownAbstractNums = new Dictionary<string, AbstractNum>();

        // This minimal numbering definition has been inspired by the documentation OfficeXMLMarkupExplained_en.docx
        // http://www.microsoft.com/downloads/details.aspx?FamilyID=6f264d0b-23e8-43fe-9f82-9ab627e5eaa3&displaylang=en
        foreach (var (listName, formatValue, text) in new[] {
            ("decimal", NumberFormatValues.Decimal, "%1."),
            ("disc", NumberFormatValues.Bullet, "•"),
            ("square", NumberFormatValues.Bullet, "▪"),
            ("circle", NumberFormatValues.Bullet, "o"),
            ("upper-alpha", NumberFormatValues.UpperLetter, "%1."),
            ("lower-alpha", NumberFormatValues.LowerLetter, "%1."),
            ("upper-roman", NumberFormatValues.UpperRoman, "%1."),
            ("lower-roman", NumberFormatValues.LowerRoman, "%1."),
        })
        {
            knownAbstractNums.Add(listName, new AbstractNum(
                new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                new Level {
                    StartNumberingValue = new() { Val = 1 },
                    NumberingFormat = new() { Val = formatValue },
                    LevelIndex = 0,
                    LevelText = new() { Val = text },
                    LevelSuffix = new() { Val = LevelSuffixValues.Tab },
                    LevelJustification = new() { Val = LevelJustificationValues.Left },
                    PreviousParagraphProperties = new() {
                        Indentation = new() { Left = "420", Hanging = "360" }
                    }
                },
                new RunProperties(
                    new RunFonts() { HighAnsi = "Arial Unicode MS" }
                )
            ) { AbstractNumDefinitionName = new() { Val = listName } });
        }

        foreach (var (listName, formatValue) in new[] {
            ("upper-greek", NumberFormatValues.UpperLetter),
            ("lower-greek", NumberFormatValues.LowerLetter),
        })
        {
            knownAbstractNums.Add(listName, new AbstractNum(
                new MultiLevelType() { Val = MultiLevelValues.SingleLevel },
                new Level {
                    StartNumberingValue = new() { Val = 1 },
                    NumberingFormat = new() { Val = formatValue },
                    LevelIndex = 0,
                    LevelText = new() { Val = "%1." },
                    PreviousParagraphProperties = new() {
                        Indentation = new() { Left = "420", Hanging = "360" }
                    },
                    NumberingSymbolRunProperties = new () {
                        RunFonts = new() { Ascii = "Symbol", Hint = FontTypeHintValues.Default }
                    }
                }
            ) { AbstractNumDefinitionName = new() { Val = listName } });
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
    /// Create a new list template 
    /// </summary>
    private int RegisterNewNumbering(ParsingContext context, ListDepth maxDepth, bool cascading = false)
    {
        Numbering numberingPart = context.MainPart.NumberingDefinitionsPart!.Numbering;

        // at this stage, we have sanitized the list style so it's safe to grab them from the predefined template lists
        var listStyles = maxDepth.Topology.Split('+');
        AbstractNum abstractNum = predefinedNumberingLists[listStyles[0]];

        abstractNum = (AbstractNum) abstractNum.CloneNode(true);
        abstractNum.AbstractNumberId = IncrementAbstractNumId(context, numberingPart);
        if (maxDepth.Depth > 1)
            abstractNum.MultiLevelType!.Val = MultiLevelValues.Multilevel;

        Level? level1 = abstractNum.GetFirstChild<Level>()!;
        // skip the first level, starts to 2
        for (int i = 1; i < maxDepth.Depth; i++)
        {
            Level level = new() {
                StartNumberingValue = new StartNumberingValue() { Val = 1 },
                NumberingFormat = new NumberingFormat() { Val = level1?.NumberingFormat?.Val },
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
                        Indentation = new Indentation() { Left = (720 * i).ToString(CultureInfo.InvariantCulture), Hanging = "360" }
                    };
            }

            abstractNum.AppendChild(level);
        }

        // this is not documented but MS Word needs that all the AbstractNum are stored consecutively.
        // Otherwise, it will apply the "NoList" style to the existing ListInstances.
        // This is the reason why I insert all the items after the last AbstractNum.
        var lastAbsNum = numberingPart.GetLastChild<AbstractNum>();
        if (lastAbsNum == null)
            numberingPart.InsertAt(abstractNum, 0);
        else
            numberingPart.InsertAfter(abstractNum, lastAbsNum);

        // register this new template
        knownAbsNumIds!.Add(maxDepth.Topology, abstractNum.AbstractNumberId);
        context.Properties("knownAbsNumIds", knownAbsNumIds);

        /*numberingPart.InsertAfter(new AbstractNum{
            AbstractNumberId = abstractNum.AbstractNumberId + 1,
            MultiLevelType = new MultiLevelType { Val = MultiLevelValues.HybridMultilevel },
            NumberingStyleLink =  new NumberingStyleLink { Val = "Harvard" }
        }, lastAbsNum);*/

            abstractNum.NumberingStyleLink =  new NumberingStyleLink { Val = "Harvard" };

        abstractNum.StyleLink = new StyleLink { Val = "Harvard" };

        context.MainPart.StyleDefinitionsPart.Styles.AddChild(
            new Style (new Name { Val = "Harvard" },
                new ParagraphProperties(
                    new NumberingProperties() { NumberingId = new NumberingId { Val = abstractNum.AbstractNumberId } }
                )) {
                Type = StyleValues.Numbering,
                StyleId = "Harvard"
            });

        return abstractNum.AbstractNumberId;// + 1;
    }
}