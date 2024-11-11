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
using System.Linq;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the ordered <c>ol</c> and unordered <c>ul</c> list.
/// </summary>
sealed class ListExpression(IHtmlElement node) : NumberingExpressionBase(node)
{
#if NET5_0_OR_GREATER
    readonly record struct ListContext(string Name, int AbsNumId, int InstanceId, int Level, DirectionMode? Dir);
#else
    readonly struct ListContext(string listName, int absNumId, int instanceId, int level, DirectionMode? dir)
    {
        public readonly string Name = listName;
        public readonly int AbsNumId = absNumId;
        public readonly int InstanceId = instanceId;
        public readonly int Level = level;
        public readonly DirectionMode? Dir = dir;
    }
#endif

    // https://www.w3schools.com/cssref/playdemo.php?filename=playcss_list-style-type
    // https://answers.microsoft.com/en-us/msoffice/forum/all/custom-list-number-style/21a54399-4404-4c37-8843-2ccaaf827485
    // Image bullet: http://officeopenxml.com/WPnumbering-imagesAsSymbol.php
    private static readonly HashSet<string> supportedListTypes = 
        ["disc", "decimal", "square", "circle",
         "lower-alpha", "upper-alpha", "lower-latin", "upper-latin",
         "lower-roman", "upper-roman",
         "decimal-tiered" /* not W3C compliant */];
    private ParagraphStyleId? listParagraphStyleId;


    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
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
        if (listContext.InstanceId == 0 || listContext.Name != listStyle)
        {
            var abstractNumId = GetOrCreateListTemplate(context, listStyle);
            listContext = ConcretiseInstance(context, abstractNumId, listStyle, listContext.Level);

            listParagraphStyleId = GetStyleIdForListItem(context.DocumentStyle, node, defaultIfEmpty: false);
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
            Paragraph p = (Paragraph) childElements.First();

            p.ParagraphProperties ??= new();
            p.ParagraphProperties.ParagraphStyleId = GetStyleIdForListItem(context.DocumentStyle, liNode);
            p.ParagraphProperties.NumberingProperties = new NumberingProperties {
                NumberingLevelReference = new() { Val = level - 1 },
                NumberingId = new() { Val = listContext.InstanceId }
            };
            if (listContext.Dir.HasValue) {
                p.ParagraphProperties.BiDi = new() {
                    Val = OnOffValue.FromBoolean(listContext.Dir == DirectionMode.Rtl)
                };
            }

            foreach (var child in childElements)
                yield return child;
        }

        context.Properties("listContext", parentContext);
    }

    /// <summary>
    /// Create a new instance of a list template.
    /// </summary>
    private ListContext ConcretiseInstance(ParsingContext context, int abstractNumId, string listStyle, int currentLevel)
    {
        ListContext listContext;

        var instanceId = GetListInstance(abstractNumId);
        int overrideLevelIndex = 0;
        var isOrderedTag = node.NodeName.Equals("ol", StringComparison.OrdinalIgnoreCase);
        var dir = node.GetTextDirection();
        if (!instanceId.HasValue || context.Converter.ContinueNumbering == false)
        {
            // create a new instance of that list template
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: context.Converter.ContinueNumbering);
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
        // be sure to restart to 1 any nested ordered list
        else if (currentLevel > 0 && isOrderedTag)
        {
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: false);
            overrideLevelIndex = currentLevel;
            listContext = new ListContext(listStyle, abstractNumId, instanceId.Value, currentLevel + 1, dir);
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

        var numbering = context.MainPart.NumberingDefinitionsPart!.Numbering;
        numbering.Append(
            new NumberingInstance(
                new AbstractNumId() { Val = abstractNumId },
                new LevelOverride(
                    new StartOverrideNumberingValue() { Val = startValue }
                ) { LevelIndex = overrideLevelIndex }
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
        string? type = styleAttributes["list-style-type"];

        if (string.IsNullOrEmpty(type) || !supportedListTypes.Contains(type!))
        {
            if (parentName != null && IsCascadingStyle(parentName))
                return parentName!;

            bool orderedList = listNode.NodeName.Equals("ol", StringComparison.OrdinalIgnoreCase);
            type = orderedList? "decimal" : "disc";
        }

        return type!;
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
}