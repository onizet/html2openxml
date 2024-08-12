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
        [TestCase("Resources.html2openxml.bmp")]
        [TestCase("Resources.html2openxml.gif")]
        [TestCase("Resources.html2openxml.jpg")]
        [TestCase("Resources.html2openxml.png")]
        [TestCase("Resources.html2openxml.emf")]
        public void GuessFormat_ReturnsImageSize(string resourceName)
        {
            using (var imageStream = ResourceHelper.GetStream(resourceName))
            {
                Size size = ImageHeader.GetDimensions(imageStream);
                Assert.Multiple(() =>
                {
                    Assert.That(size.Width, Is.EqualTo(100));
                    Assert.That(size.Height, Is.EqualTo(100));
                });
            }
        }

        [Test]
        public void AnimatedGif_ReturnsImageSize()
        {
            using (var imageStream = ResourceHelper.GetStream("Resources.stan.gif"))
            {
                Size size = ImageHeader.GetDimensions(imageStream);
                Assert.Multiple(() =>
                {
                    Assert.That(size.Width, Is.EqualTo(252));
                    Assert.That(size.Height, Is.EqualTo(318));
                });
            }
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
                Assert.Multiple(() =>
                {
                    Assert.That(size.Width, Is.EqualTo(500));
                    Assert.That(size.Height, Is.EqualTo(500));
                });
            }
        }

        [TestCase("Resources.html2openxml.bmp", ExpectedResult = ImageHeader.FileType.Bitmap)]
        [TestCase("Resources.html2openxml.gif", ExpectedResult = ImageHeader.FileType.Gif)]
        [TestCase("Resources.html2openxml.jpg", ExpectedResult = ImageHeader.FileType.Jpeg)]
        [TestCase("Resources.html2openxml.png", ExpectedResult = ImageHeader.FileType.Png)]
        public ImageHeader.FileType GuessFormat_ReturnsFileType(string resourceName)
        {
            using var imageStream = ResourceHelper.GetStream(resourceName);
            bool success = ImageHeader.TryDetectFileType(imageStream, out var guessType);

            Assert.That(success, Is.EqualTo(true));
            return guessType;
        }
    }
}