using HtmlToOpenXml.IO;
using NUnit.Framework;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests acronym, abbreviation and blockquotes.
    /// </summary>
    [TestFixture]
    public class ImageHeaderTests
    {
        [TestCase("Resources.html2openxml.bmp")]
        [TestCase("Resources.html2openxml.gif")]
        [TestCase("Resources.html2openxml.jpg")]
        [TestCase("Resources.html2openxml.png")]
        [TestCase("Resources.html2openxml.emf")]
        public void ReadSize(string resourceName)
        {
            using (var imageStream = ResourceHelper.GetStream(resourceName))
            {
                Size size = ImageHeader.GetDimensions(imageStream);
                Assert.That(size.Width, Is.EqualTo(100));
                Assert.That(size.Height, Is.EqualTo(100));
            }
        }

        [Test]
        public void ReadSizeAnimatedGif()
        {
            using (var imageStream = ResourceHelper.GetStream("Resources.stan.gif"))
            {
                Size size = ImageHeader.GetDimensions(imageStream);
                Assert.That(size.Width, Is.EqualTo(252));
                Assert.That(size.Height, Is.EqualTo(318));
            }
        }

        /// <summary>
        /// This test case cover PNG file where the dimension stands in the 2nd frame
        /// (SOF2 progressive DCT, a contrario of SOF1 baseline DCT).
        /// </summary>
        /// <remarks>https://github.com/onizet/html2openxml/issues/40</remarks>
        [Test]
        public void ReadSizePngSof2()
        {
            using (var imageStream = ResourceHelper.GetStream("Resources.lumileds.png"))
            {
                Size size = ImageHeader.GetDimensions(imageStream);
                Assert.That(size.Width, Is.EqualTo(500));
                Assert.That(size.Height, Is.EqualTo(500));
            }
        }

        [TestCase("Resources.html2openxml.bmp", IO.ImageHeader.FileType.Bitmap)]
        [TestCase("Resources.html2openxml.gif", IO.ImageHeader.FileType.Gif)]
        [TestCase("Resources.html2openxml.jpg", IO.ImageHeader.FileType.Jpeg)]
        [TestCase("Resources.html2openxml.png", IO.ImageHeader.FileType.Png)]
        public void DetectFileType(string resourceName, IO.ImageHeader.FileType type)
        {
            using (var imageStream = ResourceHelper.GetStream(resourceName))
            {
                IO.ImageHeader.FileType guessType;
                bool success = IO.ImageHeader.TryDetectFileType(imageStream, out guessType);

                Assert.That(success, Is.EqualTo(true));
                Assert.That(guessType, Is.EqualTo(type));
            }
        }
    }
}