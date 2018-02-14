using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace HtmlToOpenXml.Tests
{
    public abstract class HtmlConverterTestBase
    {
        private System.IO.MemoryStream generatedDocument;
        private WordprocessingDocument package;

        protected HtmlConverter converter;
        protected MainDocumentPart mainPart;


        [SetUp]
        public void Init ()
        {
            generatedDocument = new System.IO.MemoryStream();
            package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document);

            mainPart = package.MainDocumentPart;
            if (mainPart == null)
            {
                mainPart = package.AddMainDocumentPart();
                new Document(new Body()).Save(mainPart);
            }

            this.converter = new HtmlConverter(mainPart);
        }

        [TearDown]
        public void Close ()
        {
            package?.Dispose();
            generatedDocument?.Dispose();
        }
    }
}