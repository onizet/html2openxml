using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace HtmlToOpenXml.Tests
{
    public abstract class HtmlConverterTestBase
    {
        private MemoryStream generatedDocument = default!;
        private WordprocessingDocument package = default!;

        protected HtmlConverter converter = default!;
        protected MainDocumentPart mainPart = default!;


        [SetUp]
        public void Init ()
        {
            generatedDocument = new MemoryStream();
            package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document);

            mainPart = package.MainDocumentPart!;
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

        protected void AssertThatOpenXmlDocumentIsValid()
        {
            var validator = new OpenXmlValidator(FileFormatVersions.Office2021);
            var errors = validator.Validate(package);

            if (!errors.GetEnumerator().MoveNext())
                return;

            foreach (ValidationErrorInfo error in errors)
            {
                TestContext.Error.Write("{0}\n\t{1}\n", error.Path?.XPath, error.Description);
                if (error.Node is not null)
                    TestContext.Error.WriteLine("\n\t{0}", error.Node.OuterXml);
            }

            Assert.Fail("The document isn't conformant with Office 2021");
        }
    }
}