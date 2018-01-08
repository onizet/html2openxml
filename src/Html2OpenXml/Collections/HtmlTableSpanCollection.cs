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

namespace HtmlToOpenXml
{
    /// <summary>
    /// Typed sorted list on span in table.
    /// </summary>
    sealed class HtmlTableSpanCollection : System.Collections.ObjectModel.Collection<HtmlTableSpan>
    {
        protected override void InsertItem(int index, HtmlTableSpan item)
        {
            index = (this.Items as List<HtmlTableSpan>).BinarySearch(item);
            base.InsertItem(index < 0? ~index : index, item);
        }
    }
}