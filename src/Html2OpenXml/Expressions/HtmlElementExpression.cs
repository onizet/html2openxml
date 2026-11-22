/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using DocumentFormat.OpenXml;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Represents the base definition of the processor of an HTML tag.
/// </summary>
abstract class HtmlElementExpression : HtmlDomExpression
{
    /// <summary>
    /// Apply the style properties on the provided element.
    /// </summary>
    public abstract void CascadeStyles(OpenXmlElement element);
}
