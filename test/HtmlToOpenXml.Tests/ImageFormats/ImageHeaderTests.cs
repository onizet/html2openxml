using HtmlToOpenXml.IO;
using NUnit.Framework;

namespace HtmlToOpenXml.Tests.ImageFormats
{
    /// <summary>
    /// Tests on guessing the image format and finding its size.
    /// </summary>
    [TestFixture]
    public class ImageHeaderTests
    {
        [TestCaseSource(nameof(GuessImageSizeTestCases))]
        public void GuessFormat_ReturnsImageSize((string resourceName, Size expectedSize) td)
        {
            using (var imageStream = ResourceHelper.GetStream(td.resourceName))
            {
                Size size = ImageHeader.GetDimensions(imageStream);
                Assert.That(size, Is.EqualTo(td.expectedSize));
            }
        }

        private static IEnumerable<(string, Size)> GuessImageSizeTestCases()
        {
            yield return ("Resources.html2openxml.bmp", new Size(100, 100));
            yield return ("Resources.html2openxml.gif", new Size(100, 100));
            yield return ("Resources.html2openxml.jpg", new Size(100, 100));
            yield return ("Resources.html2openxml.png", new Size(100, 100));
            yield return ("Resources.html2openxml.emf", new Size(100, 100));
            // animated gif:
            yield return ("Resources.stan.gif", new Size(252, 318));
            yield return ("Resources.kiwi.svg", new Size(612, 502));
        }

        /// <summary>
        /// This test case cover PNG file where the dimension stands in the 2nd frame
        /// (SOF2 progressive DCT, a contrario of SOF1 baseline DCT).
        /// </summary>
        /// <remarks>https://github.com/onizet/html2openxml/issues/40</remarks>
        [Test]
        public void PngSof2_ReturnsImageSize()
        {
            using (var imageStream = ResourceHelper.GetStream("Resources.lumileds.png"))
            {
                Size size = ImageHeader.GetDimensions(imageStream);
                using (Assert.EnterMultipleScope())
                {
                    Assert.That(size.Width, Is.EqualTo(500));
                    Assert.That(size.Height, Is.EqualTo(500));
                }
            }
        }

        [TestCase("Resources.html2openxml.bmp", ExpectedResult = ImageHeader.FileType.Bitmap)]
        [TestCase("Resources.html2openxml.gif", ExpectedResult = ImageHeader.FileType.Gif)]
        [TestCase("Resources.html2openxml.jpg", ExpectedResult = ImageHeader.FileType.Jpeg)]
        [TestCase("Resources.html2openxml.png", ExpectedResult = ImageHeader.FileType.Png)]
        [TestCase("Resources.kiwi.svg", ExpectedResult = ImageHeader.FileType.Xml)]
        public ImageHeader.FileType GuessFormat_ReturnsFileType(string resourceName)
        {
            using var imageStream = ResourceHelper.GetStream(resourceName);
            bool success = ImageHeader.TryDetectFileType(imageStream, out var guessType);

            Assert.That(success, Is.True);
            return guessType;
        }

        [Test(ExpectedResult = ImageHeader.FileType.Unrecognized)]
        public ImageHeader.FileType GuessFormat_WithEmpty_ReturnsFileType()
        {
            using var memoryStream = new MemoryStream();
            bool success = ImageHeader.TryDetectFileType(memoryStream, out var guessType);

            Assert.That(success, Is.False);
            return guessType;
        }

        [TestCaseSource(nameof(ResizedImageTestCases))]
        public void KeepAspectRatio_Returns((Size actualSize, Size preferredSize, Size expectedSize) td)
        {
            var resized = ImageHeader.KeepAspectRatio(td.actualSize, td.preferredSize);
            using (Assert.EnterMultipleScope())
            {
                Assert.That(resized.Width, Is.EqualTo(td.expectedSize.Width));
                Assert.That(resized.Height, Is.EqualTo(td.expectedSize.Height));
            }
        }

        private static IEnumerable<(Size, Size, Size)> ResizedImageTestCases()
        {
            yield return (new Size(255, 0), new Size(255, 255), Size.Empty);
            yield return (new Size(255, 255), new Size(255, 255), new Size(255, 255));
            yield return (new Size(500, 255), new Size(125, 255), new Size(125, 63));
            yield return (new Size(500, 255), new Size(255, 125), new Size(255, 130));
            yield return (new Size(255, 500), new Size(255, 125), new Size(63, 125));
            yield return (new Size(500, 255), new Size(500, 750), new Size(1470, 750));
            yield return (new Size(9999, 7499), new Size(100, 75), new Size(100, 74));
            yield return (new Size(1000, 1498), new Size(0, 642), new Size(428, 642));
        }
    }
}