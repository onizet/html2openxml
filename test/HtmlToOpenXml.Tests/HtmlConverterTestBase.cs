using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace HtmlToOpenXml.Tests
{
    public abstract class HtmlConverterTestBase
    {
        private System.IO.MemoryStream _generatedDocument;
        private WordprocessingDocument _package;

        protected HtmlConverter converter;
        protected MainDocumentPart mainPart;


        [SetUp]
        public void Init ()
        {
            _generatedDocument = new System.IO.MemoryStream();
            _package = WordprocessingDocument.Create(_generatedDocument, WordprocessingDocumentType.Document);

            mainPart = _package.MainDocumentPart;
            if (mainPart == null)
            {
                mainPart = _package.AddMainDocumentPart();
                new Document(new Body()).Save(mainPart);
            }

            this.converter = new HtmlConverter(mainPart);
        }

        [TearDown]
        public void Close ()
        {
            _package?.Dispose();
            _generatedDocument?.Dispose();
        }
    }
}