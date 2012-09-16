using System;
using System.ComponentModel;
using System.Globalization;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// Represents a Html Unit (ie: 120px, 10em, ...).
	/// </summary>
	[System.Diagnostics.DebuggerDisplay("{DebuggerDisplay,nq}")]
	struct Unit
	{
		/// <summary>Represents an empty unit (not defined).</summary>
		public static readonly Unit Empty = new Unit();

		private UnitMetric type;
		private double value;
		private long valueInEmus;


		public Unit(UnitMetric type, Double value)
		{
			this.type = type;
			this.value = value;
			this.valueInEmus = ComputeInEmus(type, value);
		}

		public static Unit Parse(String str)
		{
			if (str == null) return Unit.Empty;

			str = str.Trim().ToLower(CultureInfo.InvariantCulture);
			int length = str.Length;
			int digitLength = -1;
			for (int i = 0; i < length; i++)
			{
				char ch = str[i];
				if ((ch < '0' || ch > '9') && (ch != '-' && ch != '.' && ch != ','))
					break;

				digitLength = i;
			}
			if (digitLength == -1)
			{
				// No digits in the width, we ignore this style
				return Unit.Empty;
			}

			UnitMetric type;
			if (digitLength < length - 1)
				type = ConverterUtility.ConvertToUnitMetric(str.Substring(digitLength + 1).Trim());
			else
				type = UnitMetric.Pixel;

			string v = str.Substring(0, digitLength + 1);
			double value;
			try
			{
				TypeConverter converter = new DoubleConverter();
				value = (double) converter.ConvertFromString(null, CultureInfo.InvariantCulture, v);

				if (value < Int16.MinValue || value > Int16.MaxValue)
					return Unit.Empty;
			}
			catch (FormatException)
			{
				return Unit.Empty;
			}
			catch (ArithmeticException)
			{
				return Unit.Empty;
			}

			return new Unit(type, value);
		}

		/// <summary>
		/// Gets the value expressed in the English Metrics Units.
		/// </summary>
		private static Int64 ComputeInEmus(UnitMetric type, double value)
		{
			/* Compute width and height in English Metrics Units.
			 * There are 360000 EMUs per centimeter, 914400 EMUs per inch, 12700 EMUs per point
			 * widthInEmus = widthInPixels / HorizontalResolutionInDPI * 914400
			 * heightInEmus = heightInPixels / VerticalResolutionInDPI * 914400
			 * 
			 * According to 1 px ~= 9525 EMU -> 914400 EMU per inch / 9525 EMU = 96 dpi
			 * So Word use 96 DPI printing which seems fair.
			 * http://hastobe.net/blogs/stevemorgan/archive/2008/09/15/howto-insert-an-image-into-a-word-document-and-display-it-using-openxml.aspx
			 * http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
			 *
			 * The list of units supported are explained here: http://www.w3schools.com/css/css_units.asp
			 */

			switch (type)
			{
				case UnitMetric.Percent: return 0L; // not applicable
				case UnitMetric.Emus: return (long) value;
				case UnitMetric.Inch: return (long) (value * 914400L);
				case UnitMetric.Centimeter: return (long) (value * 360000L);
				case UnitMetric.Millimeter: return (long) (value * 3600000L);
				case UnitMetric.EM:
					// well this is a rough conversion but considering 1em = 12pt (http://sureshjain.wordpress.com/2007/07/06/53/)    
					return (long) (value / 72 * 914400L * 12);
				case UnitMetric.Ex:
					return (long) (value / 72 * 914400L * 12) / 2;
				case UnitMetric.Point: return (long) (value * 12700L);
				case UnitMetric.Pica: return (long) (value / 72 * 914400L) * 12;
				case UnitMetric.Pixel: return (long) (value / 96 * 914400L);
				default: goto case UnitMetric.Pixel;
			}
		}

		//____________________________________________________________________
		//

		/// <summary>
		/// Gets the type of unit (pixel, percent, point, ...)
		/// </summary>
		public UnitMetric Type
		{
			get { return type; }
		}

		/// <summary>
		/// Gets the value of this unit.
		/// </summary>
		public Double Value
		{
			get { return value; }
		}

		/// <summary>
		/// Gets the value expressed in English Metrics Unit.
		/// </summary>
		public Int64 ValueInEmus
		{
			get { return valueInEmus; }
		}

		/// <summary>
		/// Gets the value expressed in Dxa unit.
		/// </summary>
		public Int64 ValueInDxa
		{
			get { return (long) (((double) valueInEmus / 914400L) * 20 * 72); }
		}

		/// <summary>
		/// Gets the value expressed in Pixel unit.
		/// </summary>
		public int ValueInPx
		{
			get { return (int) (type == UnitMetric.Pixel ? this.value : (float) valueInEmus / 914400L * 96); }
		}

		/// <summary>
		/// Gets the value expressed in Point unit.
		/// </summary>
		public int ValueInPoint
		{
			get { return (int) (type == UnitMetric.Point ? this.value : (float) valueInEmus / 12700L); }
		}

		/// <summary>
		/// Gets whether the unit is well formed and not empty.
		/// </summary>
		public bool IsValid
		{
			get { return this.Type != UnitMetric.Unknown; }
		}

		internal string DebuggerDisplay
		{
			get { return String.Format("{{Unit: {0} {1}}}", Value, Type); }
		}
	}
}