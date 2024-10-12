using System.Reflection;
using System.Resources;

namespace HtmlToOpenXml;

/// <summary>
/// Helper class to get chunks of OpenXml predefined style.
/// </summary>
internal class PredefinedStyles
{
    private static ResourceManager? resourceMan;

    public const string Caption = "Caption";
    public const string EndnoteText = "EndnoteText";
    public const string EndnoteReference = "EndnoteReference";
    public const string FootnoteText = "FootnoteText";
    public const string FootnoteReference = "FootnoteReference";
    public const string Heading = "Heading";
    public const string Hyperlink = "Hyperlink";
    public const string IntenseQuote = "IntenseQuote";
    public const string ListParagraph = "ListParagraph";
    public const string Quote = "Quote";
    public const string QuoteChar = "QuoteChar";
    public const string TableGrid = "TableGrid";
    public const string Header = "Header";
    public const string Footer = "Footer";
    public const string Paragraph = "Normal";



    /// <summary>
    /// Retrieves the embedded resource.
    /// </summary>
    /// <param name="styleName">The key name of the resource to find.</param>
    public static string? GetOuterXml(string styleName)
    {
        return ResourceManager.GetString(styleName);
    }

    /// <summary>
    /// Returns the cached ResourceManager instance used by this class.
    /// </summary>
    private static ResourceManager ResourceManager
    {
        get
        {
            if (resourceMan is null)
            {
                ResourceManager temp = new("HtmlToOpenXml.PredefinedStyles",
                    typeof(PredefinedStyles).GetTypeInfo().Assembly);
                resourceMan = temp;
            }
            return resourceMan;
        }
    }
}