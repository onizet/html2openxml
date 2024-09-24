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
sealed class ParsingContext(HtmlConverter converter, OpenXmlPartContainer hostingPart, IO.IImageLoader imageLoader)
{
    /// <summary>Shorthand for <see cref="Converter"/>.HtmlStyles</summary>
    public WordDocumentStyle DocumentStyle { get => Converter.HtmlStyles; }

    public HtmlConverter Converter { get; } = converter;

    public MainDocumentPart MainPart { get; } = converter.MainPart;

    public OpenXmlPartContainer HostingPart { get; } = hostingPart;

    public IO.IImageLoader ImageLoader { get; } = imageLoader;


    private HtmlElementExpression? parentExpression;
    private ParsingContext? parentContext;
    private Dictionary<string, object?> propertyBag = [];

    /// <summary>Whether the text content should preserve the line breaks.</summary>
    public bool PreserveLinebreaks { get; set; }

    /// <summary>Whether the text content should collapse the whitespaces.</summary>
    public bool CollapseWhitespaces { get; set; } = true;



    public void CascadeStyles (OpenXmlElement element)
    {
        parentExpression?.CascadeStyles(element);
        parentContext?.CascadeStyles(element);
    }

    public ParsingContext CreateChild(HtmlElementExpression expression)
    {
        var childContext = new ParsingContext(Converter, HostingPart, ImageLoader) {
            propertyBag = propertyBag,
            parentExpression = expression,
            parentContext = this
        };
        return childContext;
    }

    /// <summary>Retrieves a variable tied to the context of the parsing.</summary>
    public T? Properties<T>(string name)
        => propertyBag.TryGetValue(name, out var value)? (T?) value : default;

    /// <summary>Store a variable in the global context of the parsing.</summary>
    public void Properties(string name, object? value) => propertyBag[name] = value;
}
