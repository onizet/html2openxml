using System;
using System.ComponentModel;
using System.Globalization;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// Represents a Html Unit (ie: 120px, 10em, ...).
	/// </summary>
	struct Unit
	{
		/// <summary>Represents an empty unit (not defined).</summary>
		public static readonly Unit Empty = new Unit();

		private String type;
		private double value;
        private long valueInEmus;


		public Unit(String type, Double value)
		{
			this.type = type;
			this.value = value;
            this.valueInEmus = ComputeInEmus(type, value);
		}

		public static Unit Parse(String str)
		{
			if (str == null) return new Unit();

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
				return new Unit();
			}

			string type;
			if (digitLength < length - 1)
				type = str.Substring(digitLength + 1).Trim();
			else
				type = "px";

			string v = str.Substring(0, digitLength + 1);
			double value;
			try
			{
				TypeConverter converter = new DoubleConverter();
				value = (double) converter.ConvertFromString(null, CultureInfo.InvariantCulture, v);

				if(value < Int16.MinValue || value > Int16.MaxValue)
					return new Unit();
			}
			catch
			{
				return new Unit();
			}

			return new Unit(type, value);
		}

        /// <summary>
        /// Gets the value expressed in the English Metrics Units.
        /// </summary>
        private static Int64 ComputeInEmus(String type, double value)
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
                case "%": return 0L; // not applicable
                case "in": return (long) (value * 914400L);
                case "cm": return (long) (value * 360000L);
                case "mm": return (long) (value * 3600000L);
                case "em":
                    // well this is a rough conversion but considering 1em = 12pt (http://sureshjain.wordpress.com/2007/07/06/53/)    
                    return (long) (value / 72 * 914400L * 12);
                case "ex":
                    return (long) (value / 72 * 914400L * 12) / 2;
                case "pt": return (long) (value / 72 * 914400L);
                case "pc": return (long) (value / 72 * 914400L) * 12;
                case "px": return (long) (value / 96 * 914400L);
                default: goto case "px";
            }
        }

		//____________________________________________________________________
		//

		/// <summary>
		/// Gets the type of unit (pixel, percent, point, ...)
		/// </summary>
		public String Type 
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
            get { return (int) ((float) valueInEmus / 914400L * 96); }
        }

		public bool IsValid
		{
			get { return !String.IsNullOrEmpty(this.Type); }
		}
	}
}
