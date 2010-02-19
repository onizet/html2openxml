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
		private String type;
		private int value;


		public Unit(String type, Int32 value)
		{
			this.type = type;
			this.value = value;
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
			int value;
			try
			{
				TypeConverter converter = new SingleConverter();
				value = (int) (float) converter.ConvertFromString(null, CultureInfo.InvariantCulture, v);

				if(value < Int16.MinValue || value > Int16.MaxValue)
					return new Unit();
			}
			catch
			{
				return new Unit();
			}

			return new Unit(type, value);
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
		public Int32 Value
		{
			get { return value; }
		}

		public bool IsValid
		{
			get { return !String.IsNullOrEmpty(this.Type); }
		}
	}
}
