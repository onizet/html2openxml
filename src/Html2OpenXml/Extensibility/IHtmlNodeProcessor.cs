/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp.Dom;
using DocumentFormat.OpenXml;

namespace HtmlToOpenXml;

/// <summary>
/// Provides a custom interpretation for one or more HTML nodes.
///
/// <para>
/// Implementations can extend or override the default HtmlToOpenXml processing pipeline
/// by translating specific HTML nodes into OpenXml elements.
/// </para>
/// <para>
/// Node processors are evaluated before the built-in tag handlers. If a processor declares
/// that it can handle a node, the default HtmlToOpenXml implementation is bypassed for that node.
/// </para>
/// </summary>
public interface IHtmlNodeProcessor
{
    /// <summary>
    /// Determines whether this processor should interpret the specified node.
    /// </summary>
    /// <param name="node">The HTML node to evaluate.</param>
    /// <returns><see langword="true"/> when this processor should handle the node; otherwise, <see langword="false"/>.</returns>
    bool CanProcess (INode node);

    /// <summary>
    /// Converts the specified HTML node into OpenXml elements.
    /// </summary>
    /// <param name="context">Provides services and state information required during the interpretation process.</param>
    /// <param name="node">The node to interpret.</param>
    /// <returns>
    /// A collection of OpenXml elements representing the supplied node.
    /// An empty collection indicates that the node was intentionally consumed without producing output.
    /// </returns>
    IEnumerable<OpenXmlElement> Interpret(IHtmlNodeContext context, INode node);
}
