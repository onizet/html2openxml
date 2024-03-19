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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using HtmlToOpenXml.Expressions;

namespace HtmlToOpenXml;

/// <summary>
/// Contains information that is global to the parsing.
/// </summary>
/// <remarks>The list of paragraphs that will be returned.</remarks>
sealed class ParsingContext(HtmlConverter converter, MainDocumentPart mainPart)
{
    /// <summary>Shorthand for <see cref="Converter"/>.HtmlStyles</summary>
    public WordDocumentStyle DocumentStyle { get => Converter.HtmlStyles; }

    public HtmlConverter Converter { get; } = converter;

    public MainDocumentPart MainPart { get; } = mainPart;

    private HtmlElementExpression? parentExpression;
    private Dictionary<string, object> propertyBag = [];

    /// <summary>Whether the text content should preserver the line breaks.</summary>
    public bool PreverseLinebreaks { get; set; }


    public void CascadeStyles (OpenXmlCompositeElement element)
    {
        parentExpression?.CascadeStyles(element);

        //if (runProperties?.HasChildren != true) return;
        //TODO: DefaultRunStyle?? && DocumentStyle..DefaultRunStyle == null) return;

        //run.RunProperties = MergeStyles (run.RunProperties, runProperties);

        //if (this.DefaultRunStyle != null)
        //    run.RunProperties.RunStyle = new RunStyle() { Val = this.DefaultRunStyle };
    }


    public ParsingContext CreateChild(HtmlElementExpression expression)
    {
        var childContext = new ParsingContext(Converter, MainPart)
        {
            propertyBag = propertyBag,
            parentExpression = expression
        };
        return childContext;
    }

    //TODO: remove?
    /// <summary>
    /// Merge the properties with the tag of the previous level.
    /// </summary>
    private static T MergeStyles<T>(T? parentStyleProperties, T newStyleProperties)
        where T: OpenXmlCompositeElement, new()
    {
        if (parentStyleProperties?.HasChildren != true)
        {
            return (T) newStyleProperties.CloneNode(true);
        }
        if (newStyleProperties?.HasChildren != true)
        {
            return parentStyleProperties;
        }

        var knonwTags = new HashSet<string>();
        foreach (var prop in parentStyleProperties)
        {
            if (!knonwTags.Contains(prop.LocalName))
                knonwTags.Add(prop.LocalName);
        }

        foreach (var prop in newStyleProperties)
        {
            if (!knonwTags.Contains(prop.LocalName))
                parentStyleProperties.AddChild(prop.CloneNode(true), throwOnError: false);
        }
        return parentStyleProperties;
    }

    /// <summary>Retrieves a variable tied to the context of the parsing.</summary>
    public T? Properties<T>(string name)
        => propertyBag.TryGetValue(name, out var value)? (T) value : default;

    /// <summary>Store a variable in the global context of the parsing.</summary>
    public void Properties(string name, object value) => propertyBag[name] = value;
}
