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
		sealed class Tuple
		{
			public Table Table;
			public CellPosition CellPosition;
            public HtmlTableSpanCollection RowSpan;
		}
		private Stack<Tuple> _tables = new Stack<Tuple>(5);
		private Tuple _current;



		// IComparer<Point> Implementation

		int IComparer<CellPosition>.Compare(CellPosition x, CellPosition y)
		{
			// Only interested in the column part.
			return x.Column.CompareTo(y.Column);
		}

		public void NewContext(Table table)
		{
			if (this._current != null)
				_tables.Push(_current);

            _current = new Tuple() { Table = table, CellPosition = CellPosition.Empty, RowSpan = new HtmlTableSpanCollection() };
		}

		public void CloseContext()
		{
			if (_tables.Count > 0)
			{
				_current = _tables.Pop();
			}
			else
			{
				this._current = null;
			}
		}

		/// <summary>
		/// Tells whether the Html enumerator is currently inside any table element.
		/// </summary>
		public bool HasContext
		{
			get { return _current != null; }
		}

		/// <summary>
		/// Gets or sets the position of the current processed cell in a table.
		/// Origins is at the top left corner.
		/// </summary>
		public CellPosition CellPosition
		{
			get { return _current.CellPosition; }
			set { _current.CellPosition = value; }
		}

		/// <summary>
		/// Gets the concurrent remaining row span foreach columns (key: cell with rowSpan attribute, value: length of the span).
		/// </summary>
        public HtmlTableSpanCollection RowSpan
		{
			get { return _current.RowSpan; }
		}

		public Table CurrentTable
		{
			get { return _current.Table; }
		}
	}
}