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
    readonly record struct ListContext(string Name, int AbsNumId, int InstanceId,
        int Level, DirectionMode? Dir, long MarginLeftDxa);
    private readonly bool isOrderedTag = node.NodeName.Equals("ol", StringComparison.OrdinalIgnoreCase);
    private long listMarginLeftDxa;
    private int startValue = 1, hanging;



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
        if (!liNodes.Any()) return [];

        if (isOrderedTag && int.TryParse(node.GetAttribute("start"), out int val) && val > 1)
            startValue = val;

        var margin = styleAttributes.GetMargin();
        if (margin.Left.IsFixed && margin.Left.Value != 0)
            listMarginLeftDxa = margin.Left.ValueInDxa;

        var numberingContext = context.CreateChild(this, isStyleScoped: true);
        var liParagraphs = InterpretListItems(numberingContext, liNodes);

        if (renderAsFramed)
        {
            return [CreateFrame(liParagraphs)];
        }

        return liParagraphs;
    }

    private IEnumerable<OpenXmlElement> InterpretListItems(ParsingContext context, IEnumerable<IElement> liNodes)
    {
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

        if (listContext.InstanceId == 0 || listContext.Name != listStyle || listMarginLeftDxa > 0 || startValue != 1  )
        {
            var abstractNumId = GetOrCreateListTemplate(context, listStyle);
            listContext = ConcretiseInstance(context, abstractNumId, listStyle, listContext.Level,
                listContext.MarginLeftDxa + listMarginLeftDxa);

            listParagraphStyleId = GetStyleIdForListItem(context.DocumentStyle, node, defaultIfEmpty: false);
        }
        else
        {
            var dir = node.GetTextDirection();
            listContext = new ListContext(listContext.Name, listContext.AbsNumId, 
                // +1 because index starts on 1 and not 0
                listContext.InstanceId, listContext.Level + 1, dir ?? listContext.Dir,
                listContext.MarginLeftDxa + listMarginLeftDxa);
        }

        context.Properties("listContext", listContext);

        var leftIndentation = listContext.Level * Indentation;

        foreach (IHtmlElement liNode in liNodes.Cast<IHtmlElement>())
        {
            var expression = new BlockElementExpression(liNode);
            var childElements = expression.Interpret(context);
            if (!childElements.Any()) continue;

            // table must be aligned to the list item
            var tables = childElements.OfType<Table>();
            foreach (var table in tables)
            {
                var tableProperties = table.GetFirstChild<TableProperties>();
                if (tableProperties == null)
                    table.PrependChild(tableProperties = new());

                tableProperties.TableIndentation ??= new() { Width = leftIndentation * 2 };
                // ensure to restrain the table width to the list item
                if (tableProperties.TableWidth?.Type?.Value == TableWidthUnitValues.Pct
                    && tableProperties.TableWidth?.Width?.Value == "5000")
                {
                    tableProperties.TableWidth.Width = (5000 - leftIndentation).ToString();
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
                    NumberingLevelReference = new() { Val = listContext.Level - 1 },
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

                var liExtraMargin = ParseIndentationLeft(p.ParagraphProperties); // parsed paragraph left indentation (if any)
                p.ParagraphProperties.Indentation ??= new();
                p.ParagraphProperties.Indentation.Left = (leftIndentation * 2 + liExtraMargin - hanging).ToString(CultureInfo.InvariantCulture);
            }

            foreach (var child in childElements)
                yield return child;
        }

        // restore parent context
        context.Properties("listContext", parentContext);
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
        else if (!instanceId.HasValue || !context.Converter.ContinueNumbering || listMarginLeftDxa > 0)
        {
            // create a new instance of that list template
            bool isReusable = context.Converter.ContinueNumbering && listMarginLeftDxa == 0;
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: isReusable);
            listContext = new ListContext(listStyle, abstractNumId, instanceId.Value, newLevel, dir, listMarginLeftDxa);
        }
        else
            // if the previous element is the same list style,
            // we must restart the ordering to 0
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
        var levelOverride = CreateLevelOverride(numbering, abstractNumId, newLevel, startValue, listMarginLeftDxa,
            out var hanging);
        numbering.Append(
            new NumberingInstance(
                new AbstractNumId() { Val = abstractNumId },
                levelOverride
            )
            { NumberID = instanceId.Value });

        this.hanging = hanging;
        return listContext;
    }

    /// <summary>
    /// Builds the list instance level-override (start value and optional indent from list margin-left).
    /// </summary>
    private static LevelOverride CreateLevelOverride(Numbering numbering, int abstractNumId, int listLevel,
        int startValue, long listMarginLeftDxa, out int hanging)
    {
        int ilvl = listLevel - 1;
        var levelOverride = new LevelOverride(
            new StartOverrideNumberingValue() { Val = startValue }
        ) { LevelIndex = ilvl };

        if (listMarginLeftDxa == 0)
        {
            hanging = 0;
            return levelOverride;
        }

        long leftDxa = listLevel * Indentation * 2L + listMarginLeftDxa;
        hanging = Indentation;

        var abstractNum = numbering.Elements<AbstractNum>().First(a => a.AbstractNumberId!.Value == abstractNumId);
        var templateLevel = abstractNum.Elements<Level>().First(l => l.LevelIndex!.Value == ilvl);
        var overrideLevel = (Level)templateLevel.CloneNode(true);
        overrideLevel.PreviousParagraphProperties ??= new PreviousParagraphProperties();
        overrideLevel.PreviousParagraphProperties.Indentation = new Indentation {
            Left = leftDxa.ToString(CultureInfo.InvariantCulture),
            Hanging = hanging.ToString()
        };
        levelOverride.Level = overrideLevel;

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

    private static long ParseIndentationLeft(ParagraphProperties pPr)
    {
        if (pPr.Indentation?.Left?.Value is { } left
            && long.TryParse(left, NumberStyles.Integer, CultureInfo.InvariantCulture, out var dxa))
            return dxa;
        return 0;
    }
}
