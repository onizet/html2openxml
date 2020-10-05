using System;
using NUnit.Framework;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    using pic = DocumentFormat.OpenXml.Drawing.Pictures;

    /// <summary>
    /// Tests images.
    /// </summary>
    [TestFixture]
    public class ImgTests : HtmlConverterTestBase
    {
        [Test]
        public void ParseImg()
        {
            var elements = converter.Parse(@"<img src='https://www.w3schools.com/tags/smiley.gif' alt='Smiley face' width='42' height='42'>");
            Assert.That(elements.Count, Is.EqualTo(1));
            AssertIsImg(elements[0]);
        }

        [Test]
        public void ParseImgBorder()
        {
            var elements = converter.Parse(@"<img src='https://www.w3schools.com/tags/smiley.gif' border='1'>");
            AssertIsImg(elements[0]);
            var run = elements[0].GetFirstChild<Run>();
            RunProperties runProperties = run.GetFirstChild<RunProperties>();
            Assert.IsNotNull(runProperties);
            Assert.IsNotNull(runProperties.Border);
        }

        [Test]
        public void ParseImgManualProvisioning()
        {
            converter = new HtmlConverter(mainPart, new LocalWebRequest());

            var elements = converter.Parse(@"<img src='/img/black-dot' alt='Smiley face' width='42' height='42'>");
            Assert.That(elements.Count, Is.EqualTo(1));
            AssertIsImg(elements[0]);
        }

        private void AssertIsImg (OpenXmlCompositeElement elements)
        {
            var run = elements.GetFirstChild<Run>();
            Assert.IsNotNull(run);
            var img = run.GetFirstChild<Drawing>();
            Assert.IsNotNull(img);
            Assert.IsNotNull(img.Inline?.Graphic?.GraphicData);
            var pic = img.Inline.Graphic.GraphicData.GetFirstChild<pic.Picture>();
            Assert.IsNotNull(pic?.BlipFill?.Blip?.Embed);

            var imagePartId = pic.BlipFill.Blip.Embed.Value;
            var part = mainPart.GetPartById(imagePartId);
            Assert.That(part, Is.TypeOf(typeof(ImagePart)));
        }
    }
}