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
using System;
using AngleSharp.Dom;

namespace HtmlToOpenXml
{
    /// <summary>
    /// Helper class that provide some extension methods to AngleSharp SDK.
    /// </summary>
    static class AngleSharpExtension
    {
        /// <summary>
        /// Gets whether the given child is preceded by any list element (<c>ol</c> or <c>ul</c>).
        /// </summary>
        public static bool IsPrecededByListElement(this INode child, out IElement? precedingElement)
        {
            precedingElement = null;

            if (child.Parent == null)
                return false;

            foreach (INode childNode in child.Parent!.ChildNodes)
            {
                if (childNode == child)
                {
                    break;
                }

                if (childNode.NodeType == NodeType.Element && (
                    ((IElement) childNode).LocalName == TagNames.Ol || ((IElement) childNode).LocalName == TagNames.Ul))
                {
                    precedingElement = (IElement) childNode;
                    return true;
                }
            }

            return false;
        }
    }
}