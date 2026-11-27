/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp.Html.Dom;

namespace HtmlToOpenXml;

/// <summary>
/// Represents an HTML fragment that can be transformed, queried, and converted into OpenXml content.
///
/// <para>
/// A fragment always represent an HTML element. Non-HTML elements, such as SVG nodes, are not exposed through this abstraction.
/// </para>
/// </summary>
public sealed class HtmlFragment
{
    internal IHtmlElement Root { get; }

    internal HtmlFragment(IHtmlElement root)
    {
        Root = root;
    }

    /// <summary>
    /// Returns the first descendant fragment matching the specified selector.
    /// </summary>
    /// <param name="selectors">The group of selectors to use.</param>
    /// <returns>Returns the found fragment or <see langword="null"/> when no match is found.</returns>
    /// <remarks>
    /// Only HTML elements can be represented as <see cref="HtmlFragment"/> instances.
    /// Elements from other namespaces, such as SVG elements, are ignored and will not be returned.
    /// </remarks>
    public HtmlFragment? QuerySelector(string selectors)
    {
        var match = Root.QuerySelector(selectors);
        return match is IHtmlElement el? new HtmlFragment(el) : null;
    }

    /// <summary>
    /// Returns all descendant fragments matching the specified selector.
    /// </summary>
    /// <param name="selectors">The group of selectors to use.</param>
    /// <returns>A collection of matched fragments.</returns>
    /// <remarks>
    /// Only HTML elements are returned. Elements from other namespaces, such as SVG elements, are ignored.
    /// </remarks>
    public IReadOnlyCollection<HtmlFragment> QuerySelectoAll(string selectors)
    {
        return Root.QuerySelectorAll(selectors)
            .OfType<IHtmlElement>()
            .Select(el => new HtmlFragment(el))
            .ToArray();
    }

    /// <summary>
    /// Removes this fragment from its parent document.
    /// <para>
    /// This is useful when a fragment is rendered separately (for example as a header, footer or cover page),
    /// and should not be included in subsequent body conversion.
    /// </para>
    /// </summary>
    public void Remove()
    {
        Root.Remove();
    }

    /// <summary>
    /// Replace the content of this fragment with the specified HTML markup.
    ///
    /// <para>
    /// The supplied HTML is parsed and inserted as the new children of the current fragment.
    /// The fragment element itself is preserved
    /// </para>
    /// <para>
    /// This method modifies the loaded HTML document in-place allowing transformers and
    /// application code to reshape the document structure before OpenXml conversion begins.
    /// </para>
    /// </summary>
    /// <param name="replacementHtml">The fragment that will replace the children of this fragment.</param>
    public void ReplaceContent(string replacementHtml)
    {
        Root.InnerHtml = replacementHtml;
    }

    /// <summary>
    /// Gets or sets the inner HTML (excluding the current element) of the element.
    /// </summary>
    public string InnerHtml => Root.InnerHtml;

    /// <summary>
    /// Gets or sets the outer HTML (including the current element) of the element.
    /// </summary>
    public string OuterHtml => Root.OuterHtml;
}