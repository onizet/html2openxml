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
using System.Collections.Generic;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Base implementation for numbering list.
/// </summary>
abstract class NumberingExpressionBase(IHtmlElement node) : BlockElementExpression(node)
{
    /// <summary>Hard-coded value from Word</summary>
    public const int MaxLevel = 8;
    protected const int Indentation = 360;
    public const string HeadingNumberingName = "decimal-heading-multi";
    private static readonly IDictionary<string, AbstractNum> predefinedNumberingLists = InitKnownLists();
    /// <summary>Contains the list of templated list along with the AbstractNumbId</summary>
    private Dictionary<string, int>? knownAbsNumIds;
    /// <summary>Contains the list of numbering instance.</summary>
    private Dictionary<int, int>? knownInstanceIds;
    private bool isInitialized;


    /// <summary>
    /// Find or register an list template from the document.
    /// </summary>
    protected int GetOrCreateListTemplate(ParsingContext context, string listName)
    {
        InitNumberingIds(context);

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

    /// <summary>
    /// Find or register a list template from the document.
    /// </summary>
    protected int? GetListInstance(int abstractNumId)
    {
        if (knownInstanceIds!.TryGetValue(abstractNumId, out int instanceId))
        {
            return instanceId;
        }
        return null;
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
    protected int IncrementInstanceId(ParsingContext context, int abstractNumId, bool isReusable = true)
    {
        InitNumberingIds(context);

        var instanceId = context.Properties<int?>("listInstanceId");
        if (!instanceId.HasValue)
        {
            instanceId = 1;

            // compute the next list instance ID seed. We start at 1 because 0 has a special meaning: 
            // The w:numId can contain a value of 0, which is a special value that indicates that numbering was removed
            // at this level of the style hierarchy. While processing this markup, if the w:val='0',
            // the paragraph does not have a list item (http://msdn.microsoft.com/en-us/library/ee922775(office.14).aspx)
            var numbering = context.MainPart.NumberingDefinitionsPart!.Numbering!;
            foreach (NumberingInstance inst in numbering.Elements<NumberingInstance>())
            {
                if (inst.NumberID?.Value > instanceId) instanceId = inst.NumberID;
            }
        }
        instanceId++;
        context.Properties("listInstanceId", instanceId);

        if (isReusable)
            knownInstanceIds!.Add(abstractNumId, instanceId.Value);

        return instanceId.Value;
    }

    /// <summary>
    /// Discover the list of existing templates and instances, already registred in the document.
    /// </summary>
    private void InitNumberingIds(ParsingContext context)
    {
        if (isInitialized) return;
        knownAbsNumIds = context.Properties<Dictionary<string, int>>("knownAbsNumIds");
        knownInstanceIds = context.Properties<Dictionary<int, int>>("knownInstanceIds");
        if (knownAbsNumIds != null && knownInstanceIds != null) return;

        knownAbsNumIds = [];
        knownInstanceIds = [];
        int absNumIdRef = 0;
        NumberingDefinitionsPart numberingPart = context.MainPart.NumberingDefinitionsPart
            ?? context.MainPart.AddNewPart<NumberingDefinitionsPart>();

        if (numberingPart.Numbering == null)
        {
            new Numbering().Save(numberingPart);
        }

        var numbering = numberingPart.Numbering!;

        // The absNumIdRef Id is a required field and should be unique. We will loop through the existing Numbering definition
        // to retrieve the highest Id and reconstruct our own list definition template.
        foreach (var abs in numbering.Elements<AbstractNum>())
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
            knownInstanceIds.TryAdd(inst.AbstractNumId!.Val!.Value, inst.NumberID!.Value);
        }

        context.Properties("knownAbsNumIds", knownAbsNumIds);
        context.Properties("knownInstanceIds", knownInstanceIds);
        isInitialized = true;
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
            var abstractNum = new AbstractNum {
                AbstractNumDefinitionName = new() { Val = listName },
                MultiLevelType = new() { Val = MultiLevelValues.HybridMultilevel }
            };

            bool useSymbol = listName.EndsWith("-greek");
            for (var lvlIndex = 0; lvlIndex <= MaxLevel; lvlIndex++)
            {
                abstractNum.Append(new Level {
                    StartNumberingValue = new() { Val = 1 },
                    NumberingFormat = new() { Val = formatValue },
                    LevelIndex = lvlIndex,
                    LevelText = new() { Val = string.Format(text, lvlIndex+1) },
                    LevelJustification = new() { Val = LevelJustificationValues.Left },
                    PreviousParagraphProperties = new() {
                        Indentation = new() { Left = Indentation.ToString(), Hanging = Indentation.ToString() }
                    },
                    NumberingSymbolRunProperties = useSymbol? new () {
                        RunFonts = new() { Ascii = "Symbol", Hint = FontTypeHintValues.Default }
                    } : null
                });
            }

            knownAbstractNums.Add(listName, abstractNum);
        }

        // tiered numbering: 1, 1.1, 1.1.1
        foreach (var listName in new[] { HeadingNumberingName, "decimal-tiered" })
        {
            var abstractNum = new AbstractNum {
                AbstractNumDefinitionName = new() { Val = listName },
                MultiLevelType = new() { Val = MultiLevelValues.HybridMultilevel }
            };

            var lvlText = new System.Text.StringBuilder();
            for (var lvlIndex = 0; lvlIndex <= MaxLevel; lvlIndex++)
            {
                lvlText.AppendFormat("%{0}.", lvlIndex+1);

                abstractNum.Append(new Level {
                    StartNumberingValue = new() { Val = 1 },
                    NumberingFormat = new() { Val = NumberFormatValues.Decimal },
                    LevelIndex = lvlIndex,
                    LevelText = new() { Val = lvlText.ToString() }
                });
            }
            knownAbstractNums.Add(listName, abstractNum);
        }

        return knownAbstractNums;
    }
}