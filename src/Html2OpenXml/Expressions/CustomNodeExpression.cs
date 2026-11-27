/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Adapts an <see cref="IHtmlNodeProcessor"/> into the <see cref="HtmlDomExpression"/> pipeline.
///
/// <para>
/// This expression delegates interpretation to a registered node processor while preserving the 
/// standard HtmlToOpenXml parsing flow and context propagation.
/// </para>
/// <para>
/// The adapter allows external extensions to participate in document generation without exposing
/// internal parser implementation details.
/// </para>
/// </summary>
sealed class CustomNodeExpression(IHtmlNodeProcessor processor, INode node) : HtmlDomExpression
{
    sealed class ParsingContextAdapter(ParsingContext parsingContext) : IHtmlNodeContext
    {
        public MainDocumentPart MainPart => parsingContext.MainPart;

        public WordDocumentStyle DocumentStyle => parsingContext.DocumentStyle;

        public void CascadeStyles(OpenXmlElement element)
        {
            parsingContext.CascadeStyles(element);
        }
    }

    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        return processor.Interpret(new ParsingContextAdapter(context), node);
    }
}