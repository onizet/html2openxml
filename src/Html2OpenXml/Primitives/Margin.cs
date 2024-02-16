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

namespace HtmlToOpenXml
{
    /// <summary>
    /// Represents a Html Unit (ie: 120px, 10em, ...).
    /// </summary>
    struct Margin
    {
        private Unit[] sides;


        public Margin(Unit top, Unit right, Unit bottom, Unit left)
        {
            this.sides = [top, right, bottom, left];
        }

        /// <summary>
        /// Parse the margin style attribute.
        /// </summary>
        /// <remarks>
        /// The margin property can have from one to four values.
        /// <b>margin:25px 50px 75px 100px;</b>
        /// top margin is 25px
        /// right margin is 50px
        /// bottom margin is 75px
        /// left margin is 100px
        /// 
        /// <b>margin:25px 50px 75px;</b>
        /// top margin is 25px
        /// right and left margins are 50px
        /// bottom margin is 75px
        /// 
        /// <b>margin:25px 50px;</b>
        /// top and bottom margins are 25px
        /// right and left margins are 50px
        /// 
        /// <b>margin:25px;</b>
        /// all four margins are 25px
        /// </remarks>
        public static Margin Parse(string? str)
        {
            if (str == null) return new Margin();

            var parts = str.Split(HttpUtility.WhiteSpaces);
            switch (parts.Length)
            {
                case 1:
                    {
                        Unit all = Unit.Parse(parts[0]);
                        return new Margin(all, all, all, all);
                    }
                case 2:
                    {
                        Unit u1 = Unit.Parse(parts[0]);
                        Unit u2 = Unit.Parse(parts[1]);
                        return new Margin(u1, u2, u1, u2);
                    }
                case 3:
                    {
                        Unit u1 = Unit.Parse(parts[0]);
                        Unit u2 = Unit.Parse(parts[1]);
                        Unit u3 = Unit.Parse(parts[2]);
                        return new Margin(u1, u2, u3, u2);
                    }
                case 4:
                    {
                        Unit u1 = Unit.Parse(parts[0]);
                        Unit u2 = Unit.Parse(parts[1]);
                        Unit u3 = Unit.Parse(parts[2]);
                        Unit u4 = Unit.Parse(parts[3]);
                        return new Margin(u1, u2, u3, u4);
                    }
            }

            return new Margin();
        }

		private void EnsureSides()
		{
			if (this.sides == null) sides = new Unit[4];
		}

        //____________________________________________________________________
        //

        /// <summary>
		/// Gets or sets the unit of the bottom side.
        /// </summary>
        public Unit Bottom
        {
            get { return sides == null ? Unit.Empty : sides[2]; }
			set { EnsureSides(); sides[2] = value; }
        }

        /// <summary>
		/// Gets or sets the unit of the left side.
        /// </summary>
        public Unit Left
        {
			get { return sides == null ? Unit.Empty : sides[3]; }
			set { EnsureSides(); sides[3] = value; }
        }

        /// <summary>
		/// Gets or sets the unit of the top side.
        /// </summary>
        public Unit Top
        {
			get { return sides == null ? Unit.Empty : sides[0]; }
			set { EnsureSides(); sides[0] = value; }
        }

        /// <summary>
		/// Gets or sets the unit of the right side.
        /// </summary>
        public Unit Right
        {
			get { return sides == null ? Unit.Empty : sides[1]; }
			set { EnsureSides(); sides[1] = value; }
        }

        public bool IsValid
        {
            get => sides != null && Left.IsValid && Right.IsValid && Bottom.IsValid && Top.IsValid;
        }

		/// <summary>
		/// Gets whether at least one side has been specified.
		/// </summary>
		public bool IsEmpty
		{
			get => sides == null || !(Left.IsValid || Right.IsValid || Bottom.IsValid || Top.IsValid);
		}
    }
}