using NUnit.Framework;

namespace HtmlToOpenXml.Tests.Primitives;

/// <summary>
/// Tests inline image (base 64).
/// </summary>
[TestFixture]
public class HttpUtilityTests
{
    [TestCase("Apples &amp; Bananas <3 Here!", "Apples & Bananas <3 Here!")]
    [TestCase("Apples &not_an_entity &amp;", "Apples &not_an_entity &")]
    [TestCase("&lt;Tag&gt; &amp; \"Quotes\" here.", "<Tag> & \"Quotes\" here.", Description = "All 5 most common entities")]
    [TestCase("A line with (&amp;#38;).", "A line with (-).", Description = "Numeric entities")]
    [TestCase("&lt;a href=&quot;#page&amp;&#38;2&quot;&gt;Click Here&lt;/a&gt;", "<a href=\"#page&2\">Click Here</a>", Description = "Mixed encoding")]
    [TestCase("This text has no codes.", "This text has no codes.", Description = "No entities present, idempotency check")]
    [TestCase("One&raquo;Two", "One»Two")]
    public void Decode_ShouldSucceed(string encoded, string expected)
    {
        string actual = HtmlToOpenXml.HttpUtility.HtmlDecode(encoded);
        Assert.That(actual, Is.EqualTo(expected));
    }
}
