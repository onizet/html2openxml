using System;
using System.Collections.Frozen;
using System.Collections.Generic;

namespace HtmlToOpenXml;

/// <summary>
/// Helper class to translate a named color to its ARGB representation.
/// </summary>
partial struct HtmlColor
{
    private static readonly IReadOnlyDictionary<string, HtmlColor> namedColors = InitKnownColors();

    private static HtmlColor GetNamedColor (ReadOnlySpan<char> name)
    {
        // the longest built-in Color's name is much lower than this check, so we should not allocate here in a typical usage
        Span<char> loweredValue = name.Length <= 128 ? stackalloc char[name.Length] : new char[name.Length];

        name.ToLowerInvariant(loweredValue);

        namedColors.TryGetValue(loweredValue.ToString(), out var color);
        return color;
    }

    private static IReadOnlyDictionary<string, HtmlColor> InitKnownColors()
    {
        var colors = new Dictionary<string, HtmlColor>()
        {
            { "black", Black },
            { "white", FromArgb(255,255,255) },
            { "aliceblue", FromArgb(240, 248, 255) },
            { "lightsalmon", FromArgb(255, 160, 122) },
            { "antiquewhite", FromArgb(250, 235, 215) },
            { "lightseagreen", FromArgb(32, 178, 170) },
            { "aqua", FromArgb(0, 255, 255) },
            { "lightskyblue", FromArgb(135, 206, 250) },
            { "aquamarine", FromArgb(127, 255, 212) },
            { "lightslategray", FromArgb(119, 136, 153) },
            { "azure", FromArgb(240, 255, 255) },
            { "lightsteelblue", FromArgb(176, 196, 222) },
            { "beige", FromArgb(245, 245, 220) },
            { "lightyellow", FromArgb(255, 255, 224) },
            { "bisque", FromArgb(255, 228, 196) },
            { "lime", FromArgb(0, 255, 0) },
            { "limegreen", FromArgb(50, 205, 50) },
            { "blanchedalmond", FromArgb(255, 255, 205) },
            { "linen", FromArgb(250, 240, 230) },
            { "blue", FromArgb(0, 0, 255) },
            { "magenta", FromArgb(255, 0, 255) },
            { "blueviolet", FromArgb(138, 43, 226) },
            { "maroon", FromArgb(128, 0, 0) },
            { "brown", FromArgb(165, 42, 42) },
            { "mediumaquamarine", FromArgb(102, 205, 170) },
            { "burlywood", FromArgb(222, 184, 135) },
            { "mediumblue", FromArgb(0, 0, 205) },
            { "cadetblue", FromArgb(95, 158, 160) },
            { "mediumprchid", FromArgb(186, 85, 211) },
            { "chartreuse", FromArgb(127, 255, 0) },
            { "mediumpurple", FromArgb(147, 112, 219) },
            { "chocolate", FromArgb(210, 105, 30) },
            { "mediumseagreen", FromArgb(60, 179, 113) },
            { "coral", FromArgb(255, 127, 80) },
            { "mediumslateblue", FromArgb(123, 104, 238) },
            { "cornflowerblue", FromArgb(100, 149, 237) },
            { "mediumspringbreen", FromArgb(0, 250, 154) },
            { "cornsilk", FromArgb(255, 248, 220) },
            { "mediumturquoise", FromArgb(72, 209, 204) },
            { "crimson", FromArgb(220, 20, 60) },
            { "mediumvioletred", FromArgb(199, 21, 112) },
            { "cyan", FromArgb(0, 255, 255) },
            { "midnightblue", FromArgb(25, 25, 112) },
            { "darkblue", FromArgb(0, 0, 139) },
            { "mintcream", FromArgb(245, 255, 250) },
            { "darkcyan", FromArgb(0, 139, 139) },
            { "mistyrose", FromArgb(255, 228, 225) },
            { "darkgoldenrod", FromArgb(184, 134, 11) },
            { "moccasin", FromArgb(255, 228, 181) },
            { "darkgray", FromArgb(169, 169, 169) },
            { "navajowhite", FromArgb(255, 222, 173) },
            { "darkgreen", FromArgb(0, 100, 0) },
            { "navy", FromArgb(0, 0, 128) },
            { "darkkhaki", FromArgb(189, 183, 107) },
            { "oldlace", FromArgb(253, 245, 230) },
            { "darkmagenta", FromArgb(139, 0, 139) },
            { "olive", FromArgb(128, 128, 0) },
            { "darkolivegreen", FromArgb(85, 107, 47) },
            { "olivedrab", FromArgb(107, 142, 45) },
            { "darkorange", FromArgb(255, 140, 0) },
            { "orange", FromArgb(255, 165, 0) },
            { "darkorchid", FromArgb(153, 50, 204) },
            { "orangered", FromArgb(255, 69, 0) },
            { "darkred", FromArgb(139, 0, 0) },
            { "orchid", FromArgb(218, 112, 214) },
            { "darksalmon", FromArgb(233, 150, 122) },
            { "palegoldenrod", FromArgb(238, 232, 170) },
            { "darkseagreen", FromArgb(143, 188, 143) },
            { "palegreen", FromArgb(152, 251, 152) },
            { "darkslateblue", FromArgb(72, 61, 139) },
            { "paleturquoise", FromArgb(175, 238, 238) },
            { "darkslategray", FromArgb(40, 79, 79) },
            { "palevioletred", FromArgb(219, 112, 147) },
            { "darkturquoise", FromArgb(0, 206, 209) },
            { "papayawhip", FromArgb(255, 239, 213) },
            { "darkviolet", FromArgb(148, 0, 211) },
            { "peachpuff", FromArgb(255, 218, 155) },
            { "deeppink", FromArgb(255, 20, 147) },
            { "peru", FromArgb(205, 133, 63) },
            { "deepskyblue", FromArgb(0, 191, 255) },
            { "pink", FromArgb(255, 192, 203) },
            { "dimgray", FromArgb(105, 105, 105) },
            { "plum", FromArgb(221, 160, 221) },
            { "dodgerblue", FromArgb(30, 144, 255) },
            { "powderblue", FromArgb(176, 224, 230) },
            { "firebrick", FromArgb(178, 34, 34) },
            { "purple", FromArgb(128, 0, 128) },
            { "floralwhite", FromArgb(255, 250, 240) },
            { "red", FromArgb(255, 0, 0) },
            { "forestgreen", FromArgb(34, 139, 34) },
            { "rosybrown", FromArgb(188, 143, 143) },
            { "fuschia", FromArgb(255, 0, 255) },
            { "royalblue", FromArgb(65, 105, 225) },
            { "gainsboro", FromArgb(220, 220, 220) },
            { "saddlebrown", FromArgb(139, 69, 19) },
            { "ghostwhite", FromArgb(248, 248, 255) },
            { "salmon", FromArgb(250, 128, 114) },
            { "gold", FromArgb(255, 215, 0) },
            { "sandybrown", FromArgb(244, 164, 96) },
            { "goldenrod", FromArgb(218, 165, 32) },
            { "seagreen", FromArgb(46, 139, 87) },
            { "gray", FromArgb(128, 128, 128) },
            { "seashell", FromArgb(255, 245, 238) },
            { "green", FromArgb(0, 128, 0) },
            { "sienna", FromArgb(160, 82, 45) },
            { "greenyellow", FromArgb(173, 255, 47) },
            { "silver", FromArgb(192, 192, 192) },
            { "honeydew", FromArgb(240, 255, 240) },
            { "skyblue", FromArgb(135, 206, 235) },
            { "hotpink", FromArgb(255, 105, 180) },
            { "slateblue", FromArgb(106, 90, 205) },
            { "indianred", FromArgb(205, 92, 92) },
            { "slategray", FromArgb(112, 128, 144) },
            { "indigo", FromArgb(75, 0, 130) },
            { "snow", FromArgb(255, 250, 250) },
            { "ivory", FromArgb(255, 240, 240) },
            { "springgreen", FromArgb(0, 255, 127) },
            { "khaki", FromArgb(240, 230, 140) },
            { "steelblue", FromArgb(70, 130, 180) },
            { "lavender", FromArgb(230, 230, 250) },
            { "tan", FromArgb(210, 180, 140) },
            { "lavenderblush", FromArgb(255, 240, 245) },
            { "teal", FromArgb(0, 128, 128) },
            { "lawngreen", FromArgb(124, 252, 0) },
            { "thistle", FromArgb(216, 191, 216) },
            { "lemonchiffon", FromArgb(255, 250, 205) },
            { "tomato", FromArgb(253, 99, 71) },
            { "lightblue", FromArgb(173, 216, 230) },
            { "turquoise", FromArgb(64, 224, 208) },
            { "lightcoral", FromArgb(240, 128, 128) },
            { "violet", FromArgb(238, 130, 238) },
            { "lightcyan", FromArgb(224, 255, 255) },
            { "wheat", FromArgb(245, 222, 179) },
            { "lightgoldenrodyellow", FromArgb(250, 250, 210) },
            { "lightgreen", FromArgb(144, 238, 144) },
            { "whitesmoke", FromArgb(245, 245, 245) },
            { "lightgray", FromArgb(211, 211, 211) },
            { "yellow", FromArgb(255, 255, 0) },
            { "Lightpink", FromArgb(255, 182, 193) },
            { "yellowgreen", FromArgb(154, 205, 50) },
            { "transparent", FromArgb(0, 0, 0, 0) }
        };

        return colors.ToFrozenDictionary();
    }
}