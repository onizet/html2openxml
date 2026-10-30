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
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the ordered <c>ol</c> and unordered <c>ul</c> list.
/// </summary>
sealed class ListExpression(IHtmlElement node) : NumberingExpressionBase(node)
{
    readonly record struct ListContext(string Name, int AbsNumId, int InstanceId, int Level, DirectionMode? Dir);
    private long _listIndentationOffset;


    // https://www.w3schools.com/cssref/playdemo.php?filename=playcss_list-style-type
    // https://answers.microsoft.com/en-us/msoffice/forum/all/custom-list-number-style/21a54399-4404-4c37-8843-2ccaaf827485
    // Image bullet: http://officeopenxml.com/WPnumbering-imagesAsSymbol.php
    private static readonly HashSet<string> supportedListTypes = 
        ["disc", "decimal", "square", "circle", "dash",
         "lower-alpha", "upper-alpha", "lower-latin", "upper-latin",
         "lower-greek", "upper-greek",
         "lower-roman", "upper-roman",
         "decimal-tiered" /* not W3C compliant */];
    private ParagraphStyleId? listParagraphStyleId;


    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        ComposeStyles(context);
        var liNodes = node.Children.Where(n => n.LocalName.Equals("li", StringComparison.OrdinalIgnoreCase));
        if (!liNodes.Any()) yield break;

        // W3C requires that nested list stands below a `li` element but some editors
        // don't care to respect the standard. Let's reparent those lists
        var nestedList = node.Children.Where(n => 
            n.LocalName.Equals("ol", StringComparison.OrdinalIgnoreCase) || 
            n.LocalName.Equals("ul", StringComparison.OrdinalIgnoreCase));
        if (nestedList.Any())
        {
            foreach (var list in nestedList)
                list.PreviousElementSibling?.AppendChild(list);
        }

        var listContext = context.Properties<ListContext>("listContext");
        var parentContext = listContext;
        var listStyle = GetListName(node, listContext.Name);
        bool listMarginOnNumberingInstance = false;
        if (listContext.InstanceId == 0 || listContext.Name != listStyle || _listIndentationOffset > 0)
        {
            var abstractNumId = GetOrCreateListTemplate(context, listStyle);
            listContext = ConcretiseInstance(context, abstractNumId, listStyle, listContext.Level, _listIndentationOffset);

            listParagraphStyleId = GetStyleIdForListItem(context.DocumentStyle, node, defaultIfEmpty: false);
            listMarginOnNumberingInstance = _listIndentationOffset > 0;
        }
        else
        {
            var dir = node.GetTextDirection();
            listContext = new ListContext(listContext.Name, listContext.AbsNumId, 
                listContext.InstanceId, listContext.Level + 1, dir ?? listContext.Dir);
        }

        context.Properties("listContext", listContext);

        // +1 because index starts on 1 and not 0
        var level = Math.Min(listContext.Level, MaxLevel+1);
        foreach (IHtmlElement liNode in liNodes.Cast<IHtmlElement>())
        {
            var expression = new BlockElementExpression(liNode);
            var childElements = expression.Interpret(context);
            if (!childElements.Any()) continue;

            // table must be aligned to the list item
            var tables = childElements.OfType<Table>();
            var tableIndentation = level * Indentation;
            foreach (var table in tables)
            {
                var tableProperties = table.GetFirstChild<TableProperties>();
                if (tableProperties == null)
                    table.PrependChild(tableProperties = new());

                tableProperties.TableIndentation ??= new() { Width = tableIndentation * 2 };
                // ensure to restrain the table width to the list item
                if (tableProperties.TableWidth?.Type?.Value == TableWidthUnitValues.Pct
                    && tableProperties.TableWidth?.Width?.Value == "5000")
                {
                    tableProperties.TableWidth.Width = (5000 - tableIndentation).ToString();
                }
            }

            // ensure to filter out any non-paragraph like any nested table
            var paragraphs = childElements.OfType<Paragraph>();
            var listItemStyleId = GetStyleIdForListItem(context.DocumentStyle, liNode);

            if (paragraphs.Any())
            {
                var p = paragraphs.First();
                p.ParagraphProperties ??= new();
                p.ParagraphProperties.ParagraphStyleId = listItemStyleId;
                p.ParagraphProperties!.NumberingProperties ??= new NumberingProperties {
                    NumberingLevelReference = new() { Val = level - 1 },
                    NumberingId = new() { Val = listContext.InstanceId }
                };
                if (listContext.Dir.HasValue) {
                    p.ParagraphProperties.BiDi = new() {
                        Val = OnOffValue.FromBoolean(listContext.Dir == DirectionMode.Rtl)
                    };
                }
            }

            // any standalone paragraphs must be aligned (indented) along its current level
            foreach (var p in paragraphs.Skip(1))
            {
                // if this is a list item paragraph, skip it
                if (p.ParagraphProperties?.NumberingProperties is not null)
                    continue;

                p.ParagraphProperties ??= new();
                p.ParagraphProperties.ParagraphStyleId ??= (ParagraphStyleId?) listItemStyleId!.CloneNode(true);
                SetListItemElementParagraphIndentationLeft(p.ParagraphProperties, level, _listIndentationOffset);
            }

            foreach (var child in childElements)
                yield return child;
        }

        context.Properties("listContext", parentContext);
    }

    /// <summary>
    /// Create a new instance of a list template.
    /// </summary>
    private ListContext ConcretiseInstance(ParsingContext context, int abstractNumId, string listStyle, int currentLevel, long listMarginLeft)
    {
        ListContext listContext;

        var instanceId = GetListInstance(abstractNumId);
        int overrideLevelIndex = 0;
        var isOrderedTag = node.NodeName.Equals("ol", StringComparison.OrdinalIgnoreCase);
        var dir = node.GetTextDirection();

        // be sure to restart to 1 any nested ordered list
        if (currentLevel > 0 && isOrderedTag)
        {
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: false);
            overrideLevelIndex = currentLevel;
            listContext = new ListContext(listStyle, abstractNumId, instanceId.Value, currentLevel + 1, dir);
        }
        else if (!instanceId.HasValue || !context.Converter.ContinueNumbering || listMarginLeft > 0)
        {
            // create a new instance of that list template
            bool isReusable = context.Converter.ContinueNumbering && listMarginLeft <= 0;
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: isReusable);
            listContext = new ListContext(listStyle, abstractNumId, instanceId.Value, currentLevel + 1, dir);
        }
        else
            // if the previous element is the same list style,
            // we must restart the ordering to 0
            if (node.IsPrecededByListElement(out var precedingElement)
                && GetListName(precedingElement!) == listStyle)
        {
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: false);
            listContext =  new ListContext(listStyle, abstractNumId, instanceId.Value, 1, dir);
        }
        else
        {
            return new ListContext(listStyle, abstractNumId, instanceId.Value, currentLevel + 1, dir);
        }

        int startValue = 1;
        if (isOrderedTag)
        {
            var startAttribute = node.GetAttribute("start");
            if (startAttribute != null && int.TryParse(startAttribute, out var val) && val > 1)
                startValue = val;
        }

        var numbering = context.MainPart.NumberingDefinitionsPart!.Numbering!;
        var levelOverride = CreateLevelOverride(numbering, abstractNumId, listContext.Level, startValue, listMarginLeft);
        numbering.Append(
            new NumberingInstance(
                new AbstractNumId() { Val = abstractNumId },
                levelOverride
            )
            { NumberID = instanceId.Value });

        return listContext;
    }

    /// <summary>
    /// Builds the list instance level-override (start value and optional indent from list margin-left).
    /// </summary>
    private static LevelOverride CreateLevelOverride(Numbering numbering, int abstractNumId, int listLevel, int startValue, long listMarginLeftDxa)
    {
        int ilvl = Math.Max(0, Math.Min(listLevel, MaxLevel + 1) - 1);
        var levelOverride = new LevelOverride(
            new StartOverrideNumberingValue() { Val = startValue }
        ) { LevelIndex = ilvl };

        if (listMarginLeftDxa <= 0)
            return levelOverride;

        long leftDxa = listLevel * Indentation * 2L + listMarginLeftDxa;
        var abstractNum = numbering.Elements<AbstractNum>().First(a => a.AbstractNumberId!.Value == abstractNumId);
        var templateLevel = abstractNum.Elements<Level>().First(l => l.LevelIndex!.Value == ilvl);
        var overrideLevel = (Level)templateLevel.CloneNode(true);
        overrideLevel.PreviousParagraphProperties ??= new PreviousParagraphProperties();
        overrideLevel.PreviousParagraphProperties.Indentation = new Indentation {
            Left = leftDxa.ToString(CultureInfo.InvariantCulture),
            Hanging = Indentation.ToString()
        };
        levelOverride.Append(overrideLevel);

        return levelOverride;
    }

    /// <summary>
    /// Resolve the list style to determine which NumberList style to apply.
    /// </summary>
    private static string GetListName(IElement listNode, string? parentName = null)
    {
        var styleAttributes = listNode.GetStyles();
        bool orderedList = listNode.NodeName.Equals("ol", StringComparison.OrdinalIgnoreCase);
        string? type = styleAttributes["list-style-type"].ToString();

        // parse node "type" attribute, if inline style "list-style-type" attribute is null.
        if (string.IsNullOrEmpty(type))
            type = ListTypeToListStyleType(listNode.GetAttribute("type"), orderedList);

        if (string.IsNullOrEmpty(type) || !supportedListTypes.Contains(type!))
        {
            if (parentName != null && IsCascadingStyle(parentName))
                return parentName!;

            // If a specific list-style-type is provided for an unordered list (e.g., a custom string like "'-'"),
            // we strip the surrounding quotes to extract the raw symbol.
            // Otherwise, we fallback to the default "decimal" or "disc" styles.
            if (!orderedList && !string.IsNullOrEmpty(type))
            {
                return type!.Trim('\'', '\"');
            }
            
            return orderedList ? "decimal" : "disc";
        }

        return type!;
    }

    /// <summary>
    /// Map ordered list style attribute values to css list-style-type.
    /// Ordered lists types are "1|a|A|i|I|α" (https://w3schools.com/tags/att_ol_type.asp).
    /// Unordered lists types are "disc|circle|square" (https://developer.mozilla.org/en-US/docs/Web/HTML/Reference/Elements/ul#type).
    /// </summary>
    private static string? ListTypeToListStyleType(string? type, bool orderedList)
    {
        return (orderedList, type) switch
        {
            // un-ordered list marker style resolution
            (false, string t) when t.Equals("disc", StringComparison.OrdinalIgnoreCase) => "disc",
            (false, string t) when t.Equals("circle", StringComparison.OrdinalIgnoreCase) => "circle",
            (false, string t) when t.Equals("square", StringComparison.OrdinalIgnoreCase)  => "square",

            // ordered list marker style resolution
            (true, "1") => "decimal",
            (true, "a") => "lower-alpha",
            (true, "A") => "upper-alpha",
            (true, "i") => "lower-roman",
            (true, "I") => "upper-roman",

            _ => null
        };
    }

    /// <summary>
    /// Resolve the <see cref="ParagraphStyleId"/> of a list element node, 
    /// based on its css class if provided and if matching.
    /// </summary>
    private ParagraphStyleId? GetStyleIdForListItem(WordDocumentStyle documentStyle, IHtmlElement liNode, bool defaultIfEmpty = true)
    {
        if (listParagraphStyleId != null)
            return (ParagraphStyleId) listParagraphStyleId.Clone();

        foreach(var clsName in liNode.ClassList)
        {
            var styleId = documentStyle.GetStyle(clsName, StyleValues.Paragraph, ignoreCase: true);
            if (styleId != null)
                return new ParagraphStyleId { Val = styleId };
        }

        if (!defaultIfEmpty) return null;
        return documentStyle.GetParagraphStyle(documentStyle.DefaultStyles.ListParagraphStyle);
    }

    /// <summary>
    /// Gets whether the given style is automatically promoted to child lists.
    /// </summary>
    private static bool IsCascadingStyle(string styleName)
    {
        return styleName == "decimal-tiered";
    }

    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context); 
        var margin = styleAttributes.GetMargin();
        if (!margin.IsEmpty && margin.Left.IsFixed && margin.Left.Value > 0)
            _listIndentationOffset = margin.Left.ValueInDxa;
    }

    /// <summary>
    /// Align non-numbered paragraphs with the listnode (level + list margin + li margin).
    /// </summary>
    private static void SetListItemElementParagraphIndentationLeft(ParagraphProperties pPr, int level, long listMarginLeftDxa)
    {
        long levelLeftDxa = level * Indentation * 2;
        long liExtraDxa = ParseIndentationLeft(pPr); // parsed paragraph left indentation (if any)
        long total = levelLeftDxa + listMarginLeftDxa + liExtraDxa;

        pPr.Indentation ??= new Indentation();
        pPr.Indentation.Left = total.ToString(CultureInfo.InvariantCulture);
    }

    private static long ParseIndentationLeft(ParagraphProperties? pPr)
    {
        if (pPr?.Indentation?.Left?.Value is { } left && long.TryParse(left, NumberStyles.Integer, CultureInfo.InvariantCulture, out var dxa))
            return dxa;
        return 0;
    }
}
