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

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Represents the base definition of the processor of an HTML tag.
/// </summary>
abstract class HtmlElementExpression : HtmlDomExpression
{
    /// <summary>
    /// Apply the style properties on the provided element.
    /// </summary>
    public void CascadeStyles(OpenXmlElement element)
        => CascadeStyles(element, StyleCascade.All);

    /// <summary>
    /// Apply the style properties on the provided element.
    /// </summary>
    /// <param name="element">The OpenXml element receiving inherited properties.</param>
    /// <param name="cascade">Which inherited properties to copy. Use <see cref="StyleCascade.All"/> by default.</param>
    public abstract void CascadeStyles(OpenXmlElement element, StyleCascade cascade);
}
