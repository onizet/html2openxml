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
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Represents the collection of attributes present in the current html tag.
/// </summary>
struct HtmlAttributeCollection
{
    // we use Range to avoid string allocation
    record struct StyleAttribute(Range Name, Range Range);

    // Style key associated with a pointer to rawValue.
    private readonly MultiValueDictionary<string, StyleAttribute> attributes;
    private HtmlBorder? htmlBorder;
    private readonly string rawValue;


    private HtmlAttributeCollection(string htmlStyles)
    {
        rawValue = htmlStyles;
        attributes = new(comparer: StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Gets a value that indicates whether this collection is empty.
    /// </summary>
    public bool IsEmpty => attributes.Count == 0;

    /// <summary>
    /// Parse the <c>style</c> attribute from a HTML element.
    /// </summary>
    public static HtmlAttributeCollection ParseStyle(string? htmlStyles)
    {
        var collection = new HtmlAttributeCollection(htmlStyles!);
        if (string.IsNullOrWhiteSpace(htmlStyles)) return collection;

        var stylesSpan = htmlStyles.AsSpan();

        Span<Range> tokens = stackalloc Range[20];
        var tokenCount = stylesSpan.SplitCompositeAttribute(tokens, ';');
        if (tokenCount == 0)
            return collection;

        for (int i = 0; i < tokenCount; i++)
        {
            // Encoded ':' and ';' characters are valid for browser
            // <span style="text-decoration&#58;underline&#59;color:red">
            var segmentSpan = stylesSpan[tokens[i]];
            int index = segmentSpan.IndexOfAny('&', ':');
            int charsConsumed = 0;

            if (index == -1)
                continue; // no key:value -> ignore this style

            if ((charsConsumed = HtmlEntities.IsColon(segmentSpan.Slice(index))) > 0)
            {
                // trim() the key but keeep note of the new range
                var key = segmentSpan.Slice(0, index);
                int startTrim = 0, endTrim = index;
                while (startTrim < key.Length && char.IsWhiteSpace(key[startTrim]))
                    startTrim++;
                while (endTrim > startTrim && char.IsWhiteSpace(key[endTrim - 1]))
                    endTrim--;

                key = segmentSpan.Slice(startTrim, endTrim - startTrim);

                if (key.Length == 0) continue; // empty key

                var startIndex = index + charsConsumed;
                if (segmentSpan.Length <= startIndex + 1) continue; // empty value

                var familyKey = GetStyleFamily(key);
                var keyRng = new Range(tokens[i].Start.Value + startTrim, tokens[i].Start.Value + endTrim);
                var valueRng = new Range(tokens[i].Start.Value + startIndex, tokens[i].End);
                collection.attributes.AddWithValue(familyKey, new(keyRng, valueRng));
            }
        }

        return collection;
    }

    /// <summary>
    /// Determines the high-level, encompassing style family from a specific CSS attribute name.
    /// <para>
    /// This allows the system to correctly apply a holistic style (e.g., applying a <c>margin</c> style 
    /// when only <c>margin-top</c> was originally specified).
    /// </para>
    /// </summary>
    private static string GetStyleFamily(ReadOnlySpan<char> attributeName)
    {
        if (attributeName.StartsWith("font", StringComparison.OrdinalIgnoreCase))
            return "font";
        if (attributeName.StartsWith("padding", StringComparison.OrdinalIgnoreCase))
            return "padding";
        if (attributeName.StartsWith("margin", StringComparison.OrdinalIgnoreCase))
            return "margin";
        if (attributeName.StartsWith("border", StringComparison.OrdinalIgnoreCase))
            return "border";

        return attributeName.ToString();
    }

    /// <summary>
    /// Gets the named attribute.
    /// </summary>
    public ReadOnlySpan<char> this[string name]
    {
        get => TryGetValue(name, out var span)? span : [];
    }

    /// <summary>
    /// Determines whether the collection contains the specified key.
    /// </summary>
    public bool TryGetValue(string name, out ReadOnlySpan<char> value)
    {
        if (attributes.TryGetSingleValue(name, out var decl))
        {
            value = rawValue.AsSpan()[decl.Range].Trim();
            return true;
        }
        value = [];
        return false;
    }

    /// <summary>
    /// Efficient way to determine if a style is equals to a value.
    /// </summary>
    public bool HasStyleEqualsTo(string name, ReadOnlySpan<char> value)
    {
        if (TryGetValue(name, out var span))
        {
            return span.Equals(value, StringComparison.OrdinalIgnoreCase);
        }
        return false;
    }

    /// <summary>
    /// Gets an attribute representing a color (named color, hexadecimal or hexadecimal 
    /// without the preceding # character).
    /// </summary>
    public HtmlColor GetColor(string name)
    {
        if (TryGetValue(name, out var span))
            return HtmlColor.Parse(span);
        return HtmlColor.Empty;
    }

    /// <summary>
    /// Gets an attribute representing an unit: 120px, 10pt, 5em, 20%, ...
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="Unit.IsValid"/> property is set to false.</returns>
    public Unit GetUnit(string name, UnitMetric defaultMetric = UnitMetric.Unitless)
    {
        if (TryGetValue(name, out var span))
            return Unit.Parse(span, defaultMetric);
        return Unit.Empty;
    }

    public BoxThickness GetPadding() => GetBoxSpacing("padding");

    public BoxThickness GetMargin() => GetBoxSpacing("margin");

    /// <summary>
    /// Gets an attribute representing the 4 unit sides.
    /// If a side has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="BoxThickness.IsValid"/> property is set to false.</returns>
    private BoxThickness GetBoxSpacing(string name)
    {
        BoxThickness margin = BoxThickness.Empty;
        Unit u;

        // shortcut to avoid resolving each individual side when we know this collection is empty
        if (IsEmpty || !attributes.TryGetValue(name, out var declarations))
            return margin;

        int suffixStart = name.Length;
        foreach (var decl in declarations)
        {
            var stringSuffix = rawValue.AsSpan()[decl.Name].Slice(suffixStart);
            if (stringSuffix.IsEmpty) // the name itself, e.g., `padding`
            {
                 var m = BoxThickness.Parse(rawValue.AsSpan()[decl.Range]);
                if (m.IsValid) margin = m;
            }
            else if (stringSuffix.Equals("-block", StringComparison.OrdinalIgnoreCase))
            {
                u = Unit.Parse(rawValue.AsSpan()[decl.Range], UnitMetric.Pixel);
                if (u.IsValid) margin.Top = margin.Bottom = u;
            }
            else if (stringSuffix.Equals("-inline", StringComparison.OrdinalIgnoreCase))
            { 
                u = Unit.Parse(rawValue.AsSpan()[decl.Range], UnitMetric.Pixel);
                if (u.IsValid) margin.Left = margin.Right = u;
            }
            else if (stringSuffix.Equals("-top", StringComparison.OrdinalIgnoreCase)
                || stringSuffix.Equals("-block-start", StringComparison.OrdinalIgnoreCase))
            {
                u = Unit.Parse(rawValue.AsSpan()[decl.Range], UnitMetric.Pixel);
                if (u.IsValid) margin.Top = u;
            }
            else if (stringSuffix.Equals("-left", StringComparison.OrdinalIgnoreCase)
                || stringSuffix.Equals("-inline-start", StringComparison.OrdinalIgnoreCase))
            {
                u = Unit.Parse(rawValue.AsSpan()[decl.Range], UnitMetric.Pixel);
                if (u.IsValid) margin.Left = u;
            }
            else if (stringSuffix.Equals("-right", StringComparison.OrdinalIgnoreCase)
                || stringSuffix.Equals("-inline-end", StringComparison.OrdinalIgnoreCase))
            {
                u = Unit.Parse(rawValue.AsSpan()[decl.Range], UnitMetric.Pixel);
                if (u.IsValid) margin.Right = u;
            }
            else if (stringSuffix.Equals("-bottom", StringComparison.OrdinalIgnoreCase)
                || stringSuffix.Equals("-block-end", StringComparison.OrdinalIgnoreCase))
            {
                u = Unit.Parse(rawValue.AsSpan()[decl.Range], UnitMetric.Pixel);
                if (u.IsValid) margin.Bottom = u;
            }
        }

        return margin;
    }

    /// <summary>
    /// Gets an attribute representing the 4 border sides.
    /// If a border style/color/width has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="HtmlBorder.IsEmpty"/> property is set to false.</returns>
    public HtmlBorder GetBorders()
    {
        if (htmlBorder.HasValue)
            return htmlBorder.Value;

        // shortcut to avoid resolving each individual side when we know this collection is empty
        if (IsEmpty || !attributes.TryGetValue("border", out var declarations))
            return (htmlBorder = HtmlBorder.Empty).Value;

        HtmlBorder borders = new();
        SideBorder sb;

        foreach (var decl in declarations)
        {
            var stringSuffix = rawValue.AsSpan()[decl.Name].Slice(6 /* border */);
            if (stringSuffix.IsEmpty) // the name itself, e.g., `border`
            {
                var b = SideBorder.Parse(rawValue.AsSpan()[decl.Range]);
                if (b.IsValid) borders = new(b);
            }
            else if (stringSuffix.Equals("-block", StringComparison.OrdinalIgnoreCase))
            {
                sb = SideBorder.Parse(rawValue.AsSpan()[decl.Range]);
                if (sb.IsValid) borders.Top = borders.Bottom = sb;
            }
            else if (stringSuffix.Equals("-inline", StringComparison.OrdinalIgnoreCase))
            {
                sb = SideBorder.Parse(rawValue.AsSpan()[decl.Range]);
                if (sb.IsValid) borders.Left = borders.Right = sb;
            }
            else if (stringSuffix.StartsWith("-top", StringComparison.OrdinalIgnoreCase))
            {
                borders.Top = GetSideBorder(borders.Top, stringSuffix.Slice("-top".Length), decl);
            }
            else if (stringSuffix.StartsWith("-block-start", StringComparison.OrdinalIgnoreCase))
            {
                borders.Top = GetSideBorder(borders.Top, stringSuffix.Slice("-block-start".Length), decl);
            }
            else if (stringSuffix.StartsWith("-left", StringComparison.OrdinalIgnoreCase))
            {
                borders.Left = GetSideBorder(borders.Left, stringSuffix.Slice("-left".Length), decl);
            }
            else if (stringSuffix.StartsWith("-inline-start", StringComparison.OrdinalIgnoreCase))
            {
                borders.Left = GetSideBorder(borders.Left, stringSuffix.Slice("-inline-start".Length), decl);
            }
            else if (stringSuffix.StartsWith("-right", StringComparison.OrdinalIgnoreCase))
            {
                borders.Right = GetSideBorder(borders.Right, stringSuffix.Slice("-right".Length), decl);
            }
            else if (stringSuffix.StartsWith("-inline-end", StringComparison.OrdinalIgnoreCase))
            {
                borders.Right = GetSideBorder(borders.Right, stringSuffix.Slice("-inline-end".Length), decl);
            }
            else if (stringSuffix.StartsWith("-bottom", StringComparison.OrdinalIgnoreCase))
            {
                borders.Bottom = GetSideBorder(borders.Bottom, stringSuffix.Slice("-bottom".Length), decl);
            }
            else if (stringSuffix.StartsWith("-block-end", StringComparison.OrdinalIgnoreCase))
            {
                borders.Bottom = GetSideBorder(borders.Bottom, stringSuffix.Slice("-block-end".Length), decl);
            }
            // shorthand properties
            else if (stringSuffix.Equals("-width", StringComparison.OrdinalIgnoreCase))
            {
                borders.ParseWidths(rawValue.AsSpan()[decl.Range]);
            }
            else if (stringSuffix.Equals("-color", StringComparison.OrdinalIgnoreCase))
            {
                borders.ParseColors(rawValue.AsSpan()[decl.Range]);
            }
            else if (stringSuffix.Equals("-style", StringComparison.OrdinalIgnoreCase))
            {
                borders.ParseStyles(rawValue.AsSpan()[decl.Range]);
            }
        }

        htmlBorder = borders;
        return borders;
    }

    /// <summary>
    /// Gets an attribute representing a single border side.
    /// If a border style/color/width has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="HtmlBorder.IsEmpty"/> property is set to false.</returns>
    private SideBorder GetSideBorder(SideBorder border, ReadOnlySpan<char> name, StyleAttribute decl)
    {
        if (name.IsEmpty)
        {
            var sb = SideBorder.Parse(rawValue.AsSpan()[decl.Range]);
            if (sb.IsValid) return sb;
        }
        else if (name.Equals("-width", StringComparison.OrdinalIgnoreCase))
        {
            var u = Unit.Parse(rawValue.AsSpan()[decl.Range], UnitMetric.Unitless);
            if (u.IsFixed) return border.Apply(u);
        }
        else if (name.Equals("-color", StringComparison.OrdinalIgnoreCase))
        {
            var c = HtmlColor.Parse(rawValue.AsSpan()[decl.Range]);
            if (!c.IsEmpty) return border.Apply(c);
        }
        else if (name.Equals("-style", StringComparison.OrdinalIgnoreCase))
        {
            var s = Converter.ToBorderStyle(rawValue.AsSpan()[decl.Range]);
            if (s != BorderValues.Nil) return border.Apply(s);
        }

        return border;
    }

    /// <summary>
    /// Gets the <c>font</c> attribute and combine with the style, size and family.
    /// </summary>
    public HtmlFont GetFont()
    {
        HtmlFont font = HtmlFont.Empty;

        // shortcut to avoid resolving each individual side when we know this collection is empty
        if (IsEmpty || !attributes.TryGetValue("font", out var declarations))
            return font;

        FontStyle? fontStyle = font.Style;
        FontVariant? variant = font.Variant;
        FontWeight? weight = font.Weight;
        Unit fontSize = font.Size;
        string? family = font.Family;

        foreach (var decl in declarations)
        {
            var stringSuffix = rawValue.AsSpan()[decl.Name].Slice(4 /* font */);
            if (stringSuffix.IsEmpty)
            {
                font =  HtmlFont.Parse(rawValue.AsSpan()[decl.Range]);
                if (font.Equals(HtmlFont.Empty)) continue;

                fontStyle = font.Style;
                variant = font.Variant;
                weight = font.Weight;
                fontSize = font.Size;
                family = font.Family;
            }
            else if (stringSuffix.Equals("-style", StringComparison.OrdinalIgnoreCase))
            {
                var s = Converter.ToFontStyle(rawValue.AsSpan()[decl.Range]);
                if (s.HasValue) fontStyle = s;
            }
            else if (stringSuffix.Equals("-variant", StringComparison.OrdinalIgnoreCase))
            {
                var v = Converter.ToFontVariant(rawValue.AsSpan()[decl.Range]);
                if (v.HasValue) variant = v;
            }
            else if (stringSuffix.Equals("-weight", StringComparison.OrdinalIgnoreCase))
            {
                var w = Converter.ToFontWeight(rawValue.AsSpan()[decl.Range]);
                if (w.HasValue) weight = w;
            }
            else if (stringSuffix.Equals("-family", StringComparison.OrdinalIgnoreCase))
            {
                var f = Converter.ToFontFamily(rawValue.AsSpan()[decl.Range]);
                if (f != null) family = f;
            }
            else if (stringSuffix.Equals("-size", StringComparison.OrdinalIgnoreCase))
            {
                Unit u = Unit.Parse(rawValue.AsSpan()[decl.Range], UnitMetric.Unitless);
                if (u.IsValid) fontSize = u;
            }
        }

        return new HtmlFont(fontSize, family, fontStyle, variant, weight, Unit.Empty);
    }

    /// <summary>
    /// Gets the composite <c>text-decoration</c> style.
    /// </summary>
    public TextDecoration GetTextDecorations()
    {
        if (TryGetValue("text-decoration", out var span))
            return Converter.ToTextDecoration(span);
        return TextDecoration.None;
    }
}
