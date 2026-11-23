/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml;

/// <summary>
/// Provides access to document-level services and state during the interpretation of an HTML node.
///
/// <para>
/// This interface intentionally exposes only the information required by custom node processors and
/// shoud not be considered a complete representation of the parser state.
/// </para>
/// </summary>
public interface IHtmlNodeContext
{
    /// <summary>
    /// The mainDocumentPart where converted content will be inserted.
    /// </summary>
    MainDocumentPart MainPart { get; }

    /// <summary>
    /// Gets the style translator used for the current conversion.
    /// </summary>
    WordDocumentStyle DocumentStyle { get; }

    /// <summary>
    /// Applies all styles currently inherited from the surrounding HTML context to the specified OpenXml element.
    ///
    /// <para>
    /// Custom node processors can use this method to ensure that generated OpenXml elements participate
    /// in the same style inheritance rules as elements produced by the built-in HtmlToOpenXml processors.
    /// </para>
    /// </summary>
    /// <param name="element">The OpenXml element that should receive the inherited styles.</param>
    void CascadeStyles(OpenXmlElement element);
}
