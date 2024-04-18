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
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the ordered <c>ol</c> and unordered <c>ul</c> list.
/// </summary>
sealed class ListExpression(IHtmlElement node) : NumberingExpression(node)
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

    // https://www.w3schools.com/cssref/playdemo.php?filename=playcss_list-style-type
    // https://answers.microsoft.com/en-us/msoffice/forum/all/custom-list-number-style/21a54399-4404-4c37-8843-2ccaaf827485
    // Image bullet: http://officeopenxml.com/WPnumbering-imagesAsSymbol.php
    private static readonly HashSet<string> supportedListTypes = 
        ["disc", "decimal", "square", "circle",
         "lower-alpha", "upper-alpha", "lower-latin", "upper-latin",
         "lower-roman", "upper-roman"];
    private ParagraphStyleId? listParagraphStyleId;


    public override IEnumerable<OpenXmlCompositeElement> Interpret(ParsingContext context)
    {
        var liNodes = node.Children.Where(n => n.LocalName == "li");
        if (!liNodes.Any()) yield break;

        ListContext? parentContext = null;
        var listContext = context.Properties<ListContext>("listContext");
        var listStyle = GetListType(node);
        if (listContext.InstanceId == 0 || listContext.Name != listStyle)
        {
            parentContext = listContext;
            var abstractNumId = GetOrCreateListTemplate(context, listStyle);
            listContext = ConcretiseInstance(context, abstractNumId, listStyle, listContext.Level);

            var numbering = context.MainPart.NumberingDefinitionsPart!.Numbering;
            numbering.Append(
                new NumberingInstance(
                    new AbstractNumId() { Val = listContext.AbsNumId },
                    new LevelOverride(
                        new StartOverrideNumberingValue() { Val = 1 }
                    )
                )
                { NumberID = listContext.InstanceId });

            listParagraphStyleId = GetStyleIdForListItem(context.DocumentStyle, node, defaultIfEmpty: false);
        }
        else
        {
            parentContext = listContext;
            listContext = new ListContext(listContext.Name, listContext.AbsNumId, 
                listContext.InstanceId, listContext.Level + 1);
        }

        context.Properties("listContext", listContext);

        // +1 because index starts on 1 and not 0
        var level = Math.Min(listContext.Level, MaxLevel+1);
        foreach (IHtmlElement liNode in liNodes.Cast<IHtmlElement>())
        {
            var expression = new FlowElementExpression(liNode);
            var childElements = expression.Interpret(context);
            Paragraph p = (Paragraph) childElements.First();

            p.InsertInProperties(prop => {
                prop.ParagraphStyleId = GetStyleIdForListItem(context.DocumentStyle, liNode);
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
    /// Create a new instance of a list template.
    /// </summary>
    private ListContext ConcretiseInstance(ParsingContext context, int abstractNumId, string listStyle, int currentLevel)
    {
        var instanceId = GetListInstance(abstractNumId);
        if (!instanceId.HasValue)
        {
            // create a new instance of that list template
            instanceId = IncrementInstanceId(context, abstractNumId);
        }
        else
            // if the previous element is the same list style,
            // we must restart the ordering to 0
            if (node.PreviousElementSibling != null &&
            (node.PreviousElementSibling.LocalName == "ol" ||
             node.PreviousElementSibling.LocalName == "ul")
             && GetListType(node.PreviousElementSibling) == listStyle)
        {
            instanceId = IncrementInstanceId(context, abstractNumId, isReusable: false);
            return new ListContext(listStyle, abstractNumId, instanceId.Value, 1);
        }

        return new ListContext(listStyle, abstractNumId, instanceId.Value, currentLevel + 1);
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
}