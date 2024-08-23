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
    public abstract void CascadeStyles (OpenXmlElement element);
}
