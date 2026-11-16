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
/// Performs document-wide transformations after HTML parsing and before the OpenXml conversion begins.
///
/// <para>
/// Transformers operate on an already loaded <see cref="HtmlFragment"/> and can normalize markup, inline
/// CSS declarations, rewrite links, inject corporate branding, create bookmarks, sanitize content, 
/// or perform any DOM manipulation required by the application.
/// </para>
/// </summary>
public interface IHtmlTransformer
{
    /// <summary>
    /// Transforms the specified HTML fragment before conversion.
    /// </summary>
    /// <param name="document">The AngleSharp loaded document.</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    Task TransformAsync(AngleSharp.Dom.IDocument document, CancellationToken cancellationToken);
}