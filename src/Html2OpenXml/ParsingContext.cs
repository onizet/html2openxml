/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml.Expressions;

namespace HtmlToOpenXml;

/// <summary>
/// Contains information that is global to the parsing.
/// </summary>
sealed class ParsingContext(HtmlConverter converter, OpenXmlPartContainer hostingPart, IO.IImageLoader imageLoader)
{
    /// <summary>Shorthand for <see cref="Converter"/>.HtmlStyles</summary>
    public WordDocumentStyle DocumentStyle { get => Converter.HtmlStyles; }

    public HtmlConverter Converter { get; } = converter;

    public MainDocumentPart MainPart { get; } = converter.mainPart;

    public OpenXmlPartContainer HostingPart { get; } = hostingPart;

    public IO.IImageLoader ImageLoader { get; } = imageLoader;

    private HtmlElementExpression? parentExpression;
    private ParsingContext? parentContext;
    private Dictionary<string, object?> propertyBag = [];

    /// <summary>Whether the text content should preserve the line breaks.</summary>
    public bool PreserveLinebreaks { get; set; }

    /// <summary>Whether the text content should collapse the whitespaces.</summary>
    public bool CollapseWhitespaces { get; set; } = true;

    public void CascadeStyles(OpenXmlElement element)
    {
        parentExpression?.CascadeStyles(element);
        parentContext?.CascadeStyles(element);
    }

    /// <summary>
    /// Creates a new, nested parsing context (a child scope) by capturing the current global state. 
    /// This allows the parser to process the contained expression while maintaining a clear path for subsequent state restoration.
    /// </summary>
    /// <param name="expression">The expression owning this child context.</param>
    /// <param name="isStyleScoped">
    /// If <see langword="true"/>, this element begins a style boundary. Styles encountered here 
    /// are contained and will not propagate context changes upwards to parent nodes.
    /// This is typically used when parsing complex blocks like tables or figures.
    /// </param>
    public ParsingContext CreateChild(HtmlElementExpression expression, bool isStyleScoped = false)
    {
        return new ParsingContext(Converter, HostingPart, ImageLoader) {
            propertyBag = propertyBag,
            parentExpression = expression,
            parentContext = isStyleScoped ? null :this,
            PreserveLinebreaks = PreserveLinebreaks,
            CollapseWhitespaces = CollapseWhitespaces
        };
    }

    /// <summary>Retrieves a variable tied to the context of the parsing.</summary>
    public T? Properties<T>(string name)
        => propertyBag.TryGetValue(name, out var value)? (T?) value : default;

    /// <summary>Store a variable in the global context of the parsing.</summary>
    public void Properties(string name, object? value) => propertyBag[name] = value;
}
