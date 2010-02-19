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
		struct Tuple
		{
			public Table Table;
			public Point CellPosition;
			public SortedList<Point, Int32> RowSpan;
		}
		private Stack<Tuple> tables = new Stack<Tuple>(5);
		private Table table;
		private Point cellPosition;
		private SortedList<Point, Int32> rowSpan;


		// IComparer<Point> Implementation

		int IComparer<Point>.Compare(Point x, Point y)
		{
			// Only interested in the column part.
			return x.X.CompareTo(y.X);
		}

		public void NewContext(Table table)
		{
			if (this.table != null)
				tables.Push(new Tuple() { Table = this.table, CellPosition = this.CellPosition, RowSpan = this.rowSpan });

			this.table = table;
			rowSpan = new SortedList<Point, Int32>(this);
			this.CellPosition = Point.Empty;
		}

		public void CloseContext()
		{
			if (tables.Count > 0)
			{
				Tuple t = tables.Pop();
				this.table = t.Table;
				this.CellPosition = t.CellPosition;
				this.rowSpan = t.RowSpan;
			}
			else
			{
				this.table = null;
			}
		}

		public bool HasContext
		{
			get { return table != null; }
		}

		/// <summary>
		/// Gets or sets the position of the current processed cell in a table.
		/// Origins is at the top left corner.
		/// </summary>
		internal Point CellPosition
		{
			get { return cellPosition; }
			set { cellPosition = value; }
		}

		/// <summary>
		/// Gets the concurrent remaining row span foreach columns.
		/// </summary>
		internal SortedList<Point, Int32> RowSpan
		{
			get { return rowSpan; }
		}

		public Table CurrentTable
		{
			get { return table; }
		}
	}
}
