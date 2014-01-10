/* Copyright (C) Olivier Nizet http://html2openxml.codeplex.com - All Rights Reserved
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
using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// Holds the tables in the order we discover them (to support nested tables).
	/// </summary>
	sealed class TableContext : IComparer<Point>
	{
		sealed class Tuple
		{
			public Table Table;
			public Point CellPosition;
			public SortedList<Point, Int32> RowSpan;
		}
		private Stack<Tuple> tables = new Stack<Tuple>(5);
		private Tuple current;



		// IComparer<Point> Implementation

		int IComparer<Point>.Compare(Point x, Point y)
		{
			// Only interested in the column part.
			return x.X.CompareTo(y.X);
		}

		public void NewContext(Table table)
		{
			if (this.current != null)
				tables.Push(current);

			current = new Tuple() { Table = table, CellPosition = Point.Empty, RowSpan = new SortedList<Point, Int32>(this) };
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
		public Point CellPosition
		{
			get { return current.CellPosition; }
			set { current.CellPosition = value; }
		}

		/// <summary>
		/// Gets the concurrent remaining row span foreach columns.
		/// </summary>
		public SortedList<Point, Int32> RowSpan
		{
			get { return current.RowSpan; }
		}

		public Table CurrentTable
		{
			get { return current.Table; }
		}
	}
}
