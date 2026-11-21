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

    /// <summary>Whether the page orientation is portrait or landscape.</summary>
    public bool IsLandscape { get; set; }

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
            IsLandscape = IsLandscape,
            PreserveLinebreaks = PreserveLinebreaks,
            CollapseWhitespaces = CollapseWhitespaces
        };
    }

    /// <summary>Retrieves a variable tied to the global context of the parsing.</summary>
    public T? Properties<T>(string name)
        => propertyBag.TryGetValue(name, out var value)? (T?) value : default;

    /// <summary>Store a variable in the global context of the parsing.</summary>
    public void Properties(string name, object? value) => propertyBag[name] = value;

    /// <summary>
    /// Resolve the next available <see cref="BookmarkStart.Id"/> (they must be unique).
    /// </summary>
    public int IncrementBookmarkId()
    {
        var bookmarkRef = Properties<int?>("bookmarkRef");

        if (!bookmarkRef.HasValue)
        {
            bookmarkRef = 0;
            foreach (var b in MainPart.Document!.Body!.Descendants<BookmarkStart>())
            {
                // OpenXml SDK expose the ID as a string value but this is really an integer
                if (b.Id != null && int.TryParse(b.Id.Value, out int id) && id > bookmarkRef)
                    bookmarkRef = id;
            }
        }

        bookmarkRef++;
        Properties("bookmarkRef", bookmarkRef);
        return bookmarkRef.Value;
    }

    /// <summary>
    /// Resolve the next available <see cref="AbstractNum.AbstractNumberId"/> (they must be unique and ordered).
    /// </summary>
    public int IncrementAbstractNumId()
    {
        var absNumIdRef = Properties<int?>("absNumIdRef");
        if (!absNumIdRef.HasValue)
        {
            absNumIdRef = 0;
            // The absNumIdRef Id is a required field and should be unique. We will loop through the existing Numbering definition
            // to retrieve the highest Id and reconstruct our own list definition template.
            var numberingPart = MainPart.NumberingDefinitionsPart!.Numbering!;
            foreach (var abs in numberingPart.Elements<AbstractNum>())
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
        Properties("absNumIdRef", absNumIdRef);
        return absNumIdRef.Value;
    }
}
