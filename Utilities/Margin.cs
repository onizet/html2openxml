using System;
using System.ComponentModel;
using System.Globalization;

namespace NotesFor.HtmlToOpenXml
{
    /// <summary>
    /// Represents a Html Unit (ie: 120px, 10em, ...).
    /// </summary>
    struct Margin
    {
        static char[] whitespaces = { ' ', '\t' };

        private Unit[] sides;


        public Margin(Unit top, Unit right, Unit bottom, Unit left)
        {
            this.sides = new[] { top, right, bottom, left };
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
        public static Margin Parse(String str)
        {
            if (str == null) return new Margin();

            String[] parts = str.Split(whitespaces);
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
                        Unit u4 = Unit.Parse(parts[2]);
                        return new Margin(u1, u2, u3, u4);
                    }
            }

            return new Margin();
        }

        //____________________________________________________________________
        //

        /// <summary>
        /// Gets the unit of the bottom side.
        /// </summary>
        public Unit Bottom
        {
            get { return sides[2]; }
        }

        /// <summary>
        /// Gets the unit of the left side.
        /// </summary>
        public Unit Left
        {
            get { return sides[3]; }
        }

        /// <summary>
        /// Gets the unit of the top side.
        /// </summary>
        public Unit Top
        {
            get { return sides[0]; }
        }

        /// <summary>
        /// Gets the unit of the right side.
        /// </summary>
        public Unit Right
        {
            get { return sides[1]; }
        }

        public bool IsValid
        {
            get { return sides != null && Left.IsValid && Right.IsValid && Bottom.IsValid && Top.IsValid; }
        }
    }
}