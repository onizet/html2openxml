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
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
    /// <summary>
    /// Typed collection that holds the Style of a document and their name.
    /// OpenXml is case-sensitive but CSS is not. This collection handles both cases.
    /// </summary>
    sealed class OpenXmlDocumentStyleCollection : SortedList<String, Style>
    {
        public OpenXmlDocumentStyleCollection() : base(StringComparer.CurrentCulture)
        {
        }

        /// <summary>
        /// Gets the style associated with the specified name.
        /// </summary>
        /// <param name="name">The name whose style to get.</param>
        /// <param name="styleType">Specify the type of style seeked (Paragraph or Character).</param>
        /// <param name="style">When this method returns, the style associated with the specified name, if
        /// the key is found; otherwise, returns null. This parameter is passed uninitialized.</param>
        public bool TryGetValueIgnoreCase(String name, StyleValues styleType, out Style? style)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            // we'll use Binary Search algorithm because the collection is sorted (we inherits from SortedList)
            IList<String> keys = this.Keys;
            int low = 0, hi = keys.Count - 1, mid;

            while (low <= hi)
            {
                mid = low + (hi - low) / 2;
                // Do not use Ordinal for string comparison to avoid the '_' character not being considered (bug #13776 reported by giorand)
                int rc = String.Compare(name, keys[mid], StringComparison.CurrentCultureIgnoreCase);
                if (rc == 0)
                {
                    style = this.Values[mid];
                    Style firstFoundStyle = style;

                    // we have found the named style but maybe the style doesn't match (Paragraph is not Character)
                    for (int i = mid; i < keys.Count && !styleType.Equals(style.Type!); i++)
                    {
                        style = this.Values[i];
                        if (!name.Equals(style.StyleName!.Val, StringComparison.OrdinalIgnoreCase)) break;
                    }

                    if (!name.Equals(style.StyleName!.Val, StringComparison.OrdinalIgnoreCase))
                        style = firstFoundStyle;

                    return styleType.Equals(style.Type!);
                }
                else if (rc < 0) hi = mid - 1;
                else low = mid + 1;
            }

            style = null;
            return false;
        }
    }
}