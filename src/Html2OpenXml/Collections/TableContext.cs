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
    /// Holds the tables in the order we discover them (to support nested tables).
    /// </summary>
    sealed class TableContext : IComparer<CellPosition>
    {
        sealed class Tuple(Table table)
        {
            public Table Table = table;
            public CellPosition CellPosition = CellPosition.Empty;
            public HtmlTableSpanCollection RowSpan = [];
        }
        private Stack<Tuple> tables = new Stack<Tuple>(5);
        private Tuple? current;



        // IComparer<Point> Implementation

        int IComparer<CellPosition>.Compare(CellPosition x, CellPosition y)
        {
            // Only interested in the column part.
            return x.Column.CompareTo(y.Column);
        }

        public void NewContext(Table table)
        {
            if (this.current != null)
                tables.Push(current);

            current = new Tuple(table);
        }

        public void CloseContext()
        {
            if (tables.Count > 0)
            {
                current = tables.Pop();
            }
            else
            {
                this.current = null;
            }
        }

        /// <summary>
        /// Tells whether the Html enumerator is currently inside any table element.
        /// </summary>
        public bool HasContext
        {
            get { return current != null; }
        }

        /// <summary>
        /// Gets or sets the position of the current processed cell in a table.
        /// Origins is at the top left corner.
        /// </summary>
        public CellPosition CellPosition
        {
            get { return current?.CellPosition ?? CellPosition.Empty; }
            set {
                if (current == null) throw new InvalidOperationException();
                current.CellPosition = value;
            }
        }

        /// <summary>
        /// Gets the concurrent remaining row span foreach columns (key: cell with rowSpan attribute, value: length of the span).
        /// </summary>
        public HtmlTableSpanCollection? RowSpan
        {
            get { return current?.RowSpan; }
        }

        public Table? CurrentTable
        {
            get { return current?.Table; }
        }
    }
}