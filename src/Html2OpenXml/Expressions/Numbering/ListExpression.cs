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
    // https://developer.mozilla.org/en-US/docs/Web/CSS/Reference/Properties/list-style-type
    // https://answers.microsoft.com/en-us/msoffice/forum/all/custom-list-number-style/21a54399-4404-4c37-8843-2ccaaf827485
    // Image bullet: http://officeopenxml.com/WPnumbering-imagesAsSymbol.php
    private static readonly HashSet<string> supportedListTypes = 
        [with(StringComparer.OrdinalIgnoreCase), "disc", "decimal", "square", "circle", "dash",
         "lower-alpha", "upper-alpha", "lower-latin", "upper-latin",
         "lower-greek", "upper-greek",
         "lower-roman", "upper-roman",
         "decimal-tiered" /* not W3C compliant */];
    private const string StandaloneParagraphMarker = "list-marker";
    private ParagraphStyleId? listParagraphStyleId;
    private readonly bool isOrderedTag = node.NodeName.Equals("ol", StringComparison.OrdinalIgnoreCase);
    private long listMarginLeftDxa;
    private int startValue = 1, leftIndentation;
    private bool isNestedList;



    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        ComposeStyles(context);
        var liNodes = node.Children.Where(n => n.LocalName.Equals("li", StringComparison.OrdinalIgnoreCase));
        if (!liNodes.Any()) return [];

        if (isOrderedTag && int.TryParse(node.GetAttribute("start"), out int val) && val > 1)
            startValue = val;

        var margin = styleAttributes.GetMargin();
        if (margin.Left.IsFixed && margin.Left.Value != 0)
        {
            listMarginLeftDxa = margin.Left.ValueInDxa;
            paraProperties.Indentation = null; // clear to prevent cascading on children
        }


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

        var numberingContext = context.CreateChild(this, isStyleScoped: !isNestedList);
        var liParagraphs = InterpretListItems(numberingContext, listContext, liNodes).ToList();

        // restore parent context
        context.Properties("listContext", parentContext);

        // this marker is used to process only once any potential standalone para.
        // we remove this as DivId requires a proper registration in WebSettings.DivIds 
        if (parentContext.Level == 0)
        {
            foreach (var p in liParagraphs.OfType<Paragraph>().Where(p => p.ParagraphProperties?.DivId is not null))
                p.ParagraphProperties!.DivId!.Remove();
        }

        if (renderAsFramed)
        {
            return [CreateFrame(liParagraphs)];
        }

        return liParagraphs;
    }

    private IEnumerable<OpenXmlElement> InterpretListItems(ParsingContext context, ListContext listContext, IEnumerable<IElement> liNodes)
    {
        var listStyle = GetListName(node, listContext.Name);

        if (listContext.InstanceId == 0 || listContext.Name != listStyle || startValue != 1)
        {
            var abstractNumId = GetOrCreateListTemplate(context, listStyle);
            listContext = ConcretiseInstance(context, abstractNumId, listStyle, listContext.Level,
                listContext.MarginLeftDxa + listMarginLeftDxa);

            leftIndentation = (int) (listContext.Level * LevelIndent + listMarginLeftDxa);
            listParagraphStyleId = GetStyleIdForListItem(context.DocumentStyle, node, defaultIfEmpty: false);
        }
        else
        {
            listContext = listContext.NextLevel(node.GetTextDirection(), listMarginLeftDxa);
            leftIndentation = listContext.Level * LevelIndent;
        }

        context.Properties("listContext", listContext);
        isNestedList = listContext.Level > 1;

        foreach (IHtmlElement liNode in liNodes.Cast<IHtmlElement>())
        {
            var expression = new BlockElementExpression(liNode);
            var childElements = expression.Interpret(context);
            if (!childElements.Any()) continue;

            // table must be aligned to the list item
            var tables = childElements.OfType<Table>();
            AdjustTableIndent(tables);

            // ensure to filter out any non-paragraph like any nested table
            var paragraphs = childElements.OfType<Paragraph>();
            var listItemStyleId = GetStyleIdForListItem(context.DocumentStyle, liNode);

            var firstPara = paragraphs.FirstOrDefault();
            if (firstPara != null)
            {
                AdjustListItemParagraphIndent(firstPara, listContext, listItemStyleId);

                // any standalone paragraphs must be aligned (indented) along its current level
                AdjustStandaloneParagraphIndent(paragraphs.Skip(1), listItemStyleId);
            }

            foreach (var child in childElements)
                yield return child;
        }
    }

    private void AdjustTableIndent(IEnumerable<Table> tables)
    {
        foreach (var table in tables)
        {
            var tableProperties = table.GetFirstChild<TableProperties>();
            if (tableProperties == null)
                table.PrependChild(tableProperties = new());

            tableProperties.TableIndentation ??= new() { Width = leftIndentation };
            // ensure to restrain the table width to the list item
            if (tableProperties.TableWidth?.Type?.Value == TableWidthUnitValues.Pct
                && tableProperties.TableWidth?.Width?.Value == FullPageWidthPct)
            {
                tableProperties.TableWidth.Width = (5000 - leftIndentation).ToString();
            }
        }
    }

    private void AdjustListItemParagraphIndent(Paragraph p, ListContext listContext, ParagraphStyleId? listItemStyleId)
    {
        p.ParagraphProperties ??= new();
        p.ParagraphProperties.ParagraphStyleId = listItemStyleId;
        p.ParagraphProperties.NumberingProperties ??= new NumberingProperties {
            NumberingLevelReference = new() { Val = listContext.Level - 1 },
            NumberingId = new() { Val = listContext.InstanceId }
        };

        if (listContext.Dir.HasValue)
        {
            p.ParagraphProperties.BiDi = new() {
                Val = OnOffValue.FromBoolean(listContext.Dir == DirectionMode.Rtl)
            };
        }

        // if this <li> paragraph contains its own padding, this add to hanging
        // if the container list <ol|ul> contains its own padding, this add to left indent
        var liExtraMargin = ParseIndentationLeft(p.ParagraphProperties); // parsed paragraph left indentation (if any)
        if (liExtraMargin > 0 || listMarginLeftDxa > 0)
        {
            p.ParagraphProperties.Indentation ??= new();
            p.ParagraphProperties.Indentation.Hanging = (Hanging + liExtraMargin).ToString(CultureInfo.InvariantCulture);
            p.ParagraphProperties.Indentation.Left = leftIndentation.ToString(CultureInfo.InvariantCulture);
        }
    }

    private void AdjustStandaloneParagraphIndent(IEnumerable<Paragraph> paragraphs, ParagraphStyleId? listItemStyleId)
    {
        foreach (var p in paragraphs)
        {
            // if this is a list item paragraph, skip it
            if (p.ParagraphProperties?.NumberingProperties is not null)
            {
                continue;
            }

            p.ParagraphProperties ??= new();

            // this iteration could encounter nested paragraphs. We detect when this is DivId
            if (p.ParagraphProperties.DivId != null)
                continue;

            p.ParagraphProperties.DivId = new DivId() { Val = StandaloneParagraphMarker };
            p.ParagraphProperties.ParagraphStyleId ??= (ParagraphStyleId?) listItemStyleId!.CloneNode(true);

            // if this paragraph contains its own padding, this add to left indent
            var liExtraMargin = ParseIndentationLeft(p.ParagraphProperties); // parsed paragraph left indentation (if any)

            p.ParagraphProperties.Indentation ??= new();
            p.ParagraphProperties.Indentation.Hanging = "0";
            p.ParagraphProperties.Indentation.Left = (leftIndentation + liExtraMargin).ToString(CultureInfo.InvariantCulture);
        }
    }

    /// <summary>
    /// Create a new instance of a list template.
    /// </summary>
    private ListContext ConcretiseInstance(ParsingContext context, int abstractNumId, string listStyle,
        int currentLevel, long listMarginLeftDxa)
    {
        ListContext listContext;

        var instanceId = GetListInstance(abstractNumId);
        var dir = node.GetTextDirection();
        var newLevel = Math.Min(currentLevel + 1, MaxLevel);

        // be sure to restart to 1 any nested ordered list
        if (currentLevel > 0 && isOrderedTag)
        {
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: false);
            listContext = new ListContext(listStyle, abstractNumId, instanceId.Value, newLevel, dir, listMarginLeftDxa);
        }
        else if (!instanceId.HasValue || !context.Converter.ContinueNumbering)
        {
            // create a new instance of that list template
            bool isReusable = context.Converter.ContinueNumbering;
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: isReusable);
            listContext = new ListContext(listStyle, abstractNumId, instanceId.Value, newLevel, dir, listMarginLeftDxa);
        }
        else
            // if the previous element is the same list style,
            // we must restart the ordering to 1
            if (node.IsPrecededByListElement(out var precedingElement)
                && GetListName(precedingElement) == listStyle)
        {
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: false);
            listContext =  new ListContext(listStyle, abstractNumId, instanceId.Value, 1, dir, listMarginLeftDxa);
        }
        else
        {
            return new ListContext(listStyle, abstractNumId, instanceId.Value, newLevel, dir, listMarginLeftDxa);
        }

        var numbering = context.MainPart.NumberingDefinitionsPart!.Numbering!;
        numbering.Append(
            new NumberingInstance(
                new AbstractNumId() { Val = abstractNumId },
                new LevelOverride(
                    new StartOverrideNumberingValue() { Val = startValue }
                ) { LevelIndex = newLevel - 1 }
            )
            { NumberID = instanceId.Value });

        return listContext;
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
        foreach(var clsName in liNode.ClassList)
        {
            var styleId = documentStyle.GetStyle(clsName, StyleValues.Paragraph, ignoreCase: true);
            if (styleId != null)
                return new ParagraphStyleId { Val = styleId };
        }

        if (listParagraphStyleId != null)
            return (ParagraphStyleId) listParagraphStyleId.Clone();

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

    private static long ParseIndentationLeft(ParagraphProperties? pPr)
    {
        if (pPr?.Indentation?.Left?.Value is { } left
            && long.TryParse(left, NumberStyles.Integer, CultureInfo.InvariantCulture, out var dxa))
            return dxa;
        return 0;
    }
}
