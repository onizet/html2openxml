using System;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// Specifies the measurement values of a Html Unit.
	/// </summary>
	enum UnitMetric
	{
		Unknown,
		Percent,
		Inch,
		Centimeter,
		Millimeter,
		/// <summary>1em is equal to the current font size.</summary>
		EM,
		/// <summary>one ex is the x-height of a font (x-height is usually about half the font-size)</summary>
		Ex,
		Point,
		Pica,
		Pixel,

		// this value is not parsed but can be used internally
		Emus
	}
}