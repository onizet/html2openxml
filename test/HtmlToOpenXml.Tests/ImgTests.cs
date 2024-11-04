using NUnit.Framework;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Moq;

namespace HtmlToOpenXml.Tests
{
    using pic = DocumentFormat.OpenXml.Drawing.Pictures;
    using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

    /// <summary>
    /// Tests images.
    /// </summary>
    [TestFixture]
    public class ImgTests : HtmlConverterTestBase
    {
        [TestCase("https://www.w3schools.com/tags/smiley.gif", "image/gif")]
        [TestCase("https://dev.w3.org/SVG/tools/svgweb/samples/svg-files/helloworld.svg", "image/svg+xml")]
        public void AbsoluteUri_ReturnsDrawing_WithDownloadedData(string imageUri, string contentType)
        {
            var elements = converter.Parse(@$"<img src='{imageUri}' alt='Smiley face' width='42' height='42'>");
            Assert.That(elements, Has.Count.EqualTo(1));
            var (_, imagePart) = AssertIsImg(mainPart, elements[0]);
            Assert.That(imagePart.ContentType, Is.EqualTo(contentType));
        }

        [Test]
        public void DataUri_ReturnsDrawing_WithDecryptedData()
        {
            var elements = converter.Parse(@"<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==' alt='Smiley face' width='42' height='42'>");
            Assert.That(elements, Has.Count.EqualTo(1));
            AssertIsImg(mainPart, elements[0]);
        }

        [Test]
        public void WithBorder_ReturnsRunWithBorder()
        {
            var elements = converter.Parse(@"<img src='https://www.w3schools.com/tags/smiley.gif' border='1'>");
            AssertIsImg(mainPart, elements[0]);
            var run = elements[0].GetFirstChild<Run>();
            var runProperties = run?.GetFirstChild<RunProperties>();
            Assert.That(runProperties, Is.Not.Null);
            Assert.That(runProperties.Border, Is.Not.Null);
        }

        [Test]
        public void ManualProvisioning_ReturnsDrawing_WithProvidedData()
        {
            var webRequest = new Mock<IO.IWebRequest>();
            webRequest.Setup(x => x.FetchAsync(It.IsAny<Uri>(), It.IsAny<CancellationToken>()))
                .Returns(Task.FromResult<IO.Resource?>(new() {
                    Content = new MemoryStream(Convert.FromBase64String(@"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")),
                    StatusCode = System.Net.HttpStatusCode.OK
                }));
            converter = new HtmlConverter(mainPart, webRequest.Object);

            var elements = converter.Parse(@"<img src='/img/black-dot' alt='Smiley face' width='42' height='42'>");
            Assert.That(elements, Has.Count.EqualTo(1));
            AssertIsImg(mainPart, elements[0]);
        }

        [TestCase("<img src='tcp://192.168.0.1:53/attach.jpg'>", Description = "Unsupported protocol")]
        [TestCase("<img src='/attach.jpg'>", Description = "Relative url without providing BaseImagerUri")]
        public void IgnoreImage_ShouldBeIgnored(string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Is.Empty);
        }

        [Test]
        public void ManualProvisioning_WithNoContent_ShouldBeIgnored()
        {
            var webRequest = new Mock<IO.IWebRequest>();
            webRequest.Setup(x => x.FetchAsync(It.IsAny<Uri>(), It.IsAny<CancellationToken>()))
                .Returns(Task.FromResult<IO.Resource?>(null));
            converter = new HtmlConverter(mainPart, webRequest.Object);

            var elements = converter.Parse(@$"<img src='/images/{Guid.NewGuid()}.png'>");
            Assert.That(elements, Is.Empty);
        }

        [Test(Description = "Reading image from a local base image url.")]
        public async Task FileSystem_LocalImage_WithBaseUri_ShouldSucceed()
        {
            string baseUri = TestContext.CurrentContext.WorkDirectory;

            using var resourceStream = ResourceHelper.GetStream("Resources.html2openxml.jpg");
            using (var fileStream = File.OpenWrite(Path.Combine(baseUri, "html2openxml.jpg")))
                await resourceStream.CopyToAsync(fileStream);

            converter = new(mainPart, new IO.DefaultWebRequest { BaseImageUrl = new Uri(baseUri) });

            var elements = await converter.ParseAsync("<img src='html2openxml.jpg'>");
            Assert.That(elements.Count(), Is.EqualTo(1));
            AssertIsImg(mainPart, elements.First());
        }

        [Test(Description = "Reading local file containing a space in the name")]
        public async Task FileSystem_LocalImage_WithSpaceInName_ShouldSucceed()
        {
            string filepath = Path.Combine(TestContext.CurrentContext.WorkDirectory, @"html2openxml copy.jpg");

            using var resourceStream = ResourceHelper.GetStream("Resources.html2openxml.jpg");
            using (var fileStream = File.OpenWrite(filepath))
                await resourceStream.CopyToAsync(fileStream);

            var localUri = "file:///" + filepath.TrimStart('/').Replace(" ", "%20");
            var elements = await converter.ParseAsync($"<img src='{localUri}'>");
            Assert.That(elements.Count(), Is.EqualTo(1));
            AssertIsImg(mainPart, elements.First());
        }

        [Test]
        public void SvgNode_ReturnsImage()
        {
            var elements = converter.Parse(ResourceHelper.GetString("Resources.kiwi.svg"));
            Assert.That(elements, Has.Count.EqualTo(1));
            var (drawing, imagePart) = AssertIsImg(mainPart, elements[0]);
            Assert.That(drawing.Inline!.DocProperties?.Name?.Value, Is.EqualTo("Illustration of a Kiwi"));
            Assert.That(drawing.Inline!.DocProperties?.Description?.Value, Is.EqualTo("Kiwi (/ˈkiːwiː/ KEE-wee)[4] are flightless birds endemic to New Zealand of the order Apterygiformes."));
            Assert.That(imagePart.ContentType, Is.EqualTo("image/svg+xml"));
        }

        [Test(Description = "Reading local file containing a space in the name")]
        public async Task RemoteImage_WithBaseUri_ShouldSucceed()
        {
            converter = new HtmlConverter(mainPart, new IO.DefaultWebRequest() { 
                BaseImageUrl = new Uri("http://github.com/onizet/html2openxml")
            });
            var elements = await converter.ParseAsync($"<img src='/blob/dev/icon.png'>");
            Assert.That(elements, Is.Not.Empty);
            AssertIsImg(mainPart, elements.First());
        }

        [Test(Description = "Image ID must be unique, amongst header, body and footer parts")]
        public async Task ImageIds_IsUniqueAcrossPackagingParts()
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithImgHeaderFooter.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;

            var beforeMaxDocPropId = new[] {
                mainPart.Document.Body!.Descendants<wp.DocProperties>(),
                mainPart.HeaderParts.SelectMany(x => x.Header.Descendants<wp.DocProperties>()),
                mainPart.FooterParts.SelectMany(x => x.Footer.Descendants<wp.DocProperties>())
            }.SelectMany(x => x).MaxBy(x => x.Id?.Value ?? 0)?.Id?.Value;
            Assert.That(beforeMaxDocPropId, Is.Not.Null);

            HtmlConverter converter = new(mainPart);
            await converter.ParseBody("<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==' width='42' height='42'>");
            mainPart.Document.Save();

            var img = mainPart.Document.Body!.Descendants<Drawing>().FirstOrDefault();
            Assert.That(img, Is.Not.Null);
            Assert.That(img.Inline?.DocProperties?.Id?.Value,
                Is.GreaterThan(beforeMaxDocPropId),
                "New image id is incremented considering existing images in header, body and footer");
        }

        [GenericTestCase(typeof(HeaderPart), Description = "Incomplete header or footer definition must be skipped #159")]
        [GenericTestCase(typeof(FooterPart))]
        public void WithIncompleteHeader_ShouldNotThrow<T>() where T : OpenXmlPart, IFixedContentTypePart
        {
            using var generatedDocument = new MemoryStream();
            using (var buffer = ResourceHelper.GetStream("Resources.DocWithImgHeaderFooter.docx"))
                buffer.CopyTo(generatedDocument);

            generatedDocument.Position = 0L;
            using WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);
            MainDocumentPart mainPart = package.MainDocumentPart!;
            mainPart.AddNewPart<T>(); // this code is incomplete as it's missing the header content

            HtmlConverter converter = new(mainPart);
            Assert.DoesNotThrowAsync(async () =>
                await converter.ParseBody("<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==' width='42' height='42'>"));
        }

        [TestCase(typeof(HeaderPart))]
        [TestCase(typeof(FooterPart))]
        [TestCase(typeof(MainDocumentPart))]
        public async Task ParseIntoDocumentPart_ReturnsImageParentedToPart (Type openXmlPartType)
        {
            string html = @"<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==' alt='Smiley face' width='42' height='42'>";
            OpenXmlElement host;
            OpenXmlPartContainer container;

            if (openXmlPartType == typeof(HeaderPart))
            {
                await converter.ParseHeader(html);
                container = mainPart.HeaderParts.First();
                host = mainPart.HeaderParts.First().Header;
            }
            else if (openXmlPartType == typeof(FooterPart))
            {
                await converter.ParseFooter(html);
                container = mainPart.FooterParts.First();
                host = mainPart.FooterParts.First().Footer;
            }
            else if (openXmlPartType == typeof(MainDocumentPart))
            {
                await converter.ParseBody(html);
                container = mainPart;
                host = mainPart.Document.Body!;
            }
            else
            {
                throw new NotSupportedException($"Test case not supported for {openXmlPartType.FullName}");
            }

            Assert.That(host.ChildElements, Has.Count.EqualTo(1));
            var p = host.ChildElements.FirstOrDefault(c => c is Paragraph);
            Assert.That(p, Is.Not.Null);
            AssertIsImg(container, p);
            AssertThatOpenXmlDocumentIsValid();
        }

        [TestCase("block", ExpectedResult = true)]
        [TestCase("flex", ExpectedResult = true)]
        [TestCase("inline", ExpectedResult = false)]
        public bool CenterImg_ReturnsFramedImg(string displayMode)
        {
            var elements = converter.Parse($@"<img style=""display: {displayMode}; margin-left: auto; margin-right: auto;""
                src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg=="" width=""32"" height=""32"">");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0], Is.TypeOf<Paragraph>());
            AssertIsImg(mainPart, elements[0]);
            return elements[0].GetFirstChild<ParagraphProperties>()?.
                Justification?.Val?.Value == JustificationValues.Center;
        }

        private static (Drawing, ImagePart) AssertIsImg (OpenXmlPartContainer container, OpenXmlElement paragraph)
        {
            var run = paragraph.GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);
            var drawing = run.GetFirstChild<Drawing>();
            Assert.That(drawing, Is.Not.Null);
            Assert.That(drawing.Inline?.Graphic?.GraphicData, Is.Not.Null);
            var pic = drawing.Inline.Graphic.GraphicData.GetFirstChild<pic.Picture>();
            Assert.That(pic?.BlipFill?.Blip?.Embed, Is.Not.Null);

            var imagePartId = pic.BlipFill.Blip.Embed.Value;
            Assert.That(imagePartId, Is.Not.Null);
            var imagePart = container.GetPartById(imagePartId);
            Assert.That(imagePart, Is.TypeOf(typeof(ImagePart)));
            return (drawing, (ImagePart) imagePart);
        }
    }
}