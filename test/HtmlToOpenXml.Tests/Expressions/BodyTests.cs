using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests.Expressions;

/// <summary>
/// Tests on <c>body</c> elements.
/// </summary>
[TestFixture]
public class BodyTests : HtmlConverterTestBase
{
    [TestCase("rtl", ExpectedResult = true)]
    [TestCase("ltr", ExpectedResult = false)]
    [TestCase("", ExpectedResult = null)]
    public bool? WithRtl_ReturnsBidi_DocumentScoped(string dir)
    {
        var elements = converter.Parse($@"<body dir='{dir}'>Lorem</body>");
        Assert.That(elements, Has.Count.EqualTo(1));
        Assert.That(elements, Has.All.TypeOf<Paragraph>());

        var bidi = mainPart.Document!.Body!.GetFirstChild<SectionProperties>()?.GetFirstChild<BiDi>();
        return bidi?.Val?.Value;
    }
}
