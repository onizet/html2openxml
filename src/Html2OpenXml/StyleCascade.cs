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

namespace HtmlToOpenXml;

/// <summary>
/// Selects which inherited style properties are copied onto descendant OpenXml elements.
/// Add a property here when a new kind of cascade must be suppressed without changing call-site signatures.
/// </summary>
record struct StyleCascade
{
    /// <summary>Copy ancestor run <c>Shading</c> (background-color).</summary>
    public bool RunShading { get; set; } = true;

    public StyleCascade() { }

    /// <summary>Copy every supported inherited style.</summary>
    public static StyleCascade All { get; } = new();
}
