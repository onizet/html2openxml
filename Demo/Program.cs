using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using NotesFor.HtmlToOpenXml;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            const string filename = "test.docx";
            string html = Demo.Properties.Resources.CompleteRunTest;
            if (File.Exists(filename)) File.Delete(filename);

            using (MemoryStream generatedDocument = new MemoryStream())
            {
                // Uncomment and comment the second using() to open an existing template document
                // instead of creating it from scratch.

                byte[] data = Demo.Properties.Resources.template;
                generatedDocument.Write(data, 0, data.Length);
                generatedDocument.Position = 0L;
                using (WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true))
                //using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }

                    HtmlConverter converter = new HtmlConverter(mainPart);
                    //converter.WebProxy.Credentials = new System.Net.NetworkCredential("nizeto", "****", "domain");
                    //converter.WebProxy.Proxy = new System.Net.WebProxy("proxy01:8080");
                    converter.ImageProcessing = ImageProcessing.ManualProvisioning;
                    converter.ProvisionImage += OnProvisionImage;

                    Body body = mainPart.Document.Body;

                    converter.ParseHtml(html);
                    mainPart.Document.Save();

                    AssertThatOpenXmlDocumentIsValid(package);
                }

                File.WriteAllBytes(filename, generatedDocument.ToArray());
            }

            System.Diagnostics.Process.Start(filename);
        }

        static void OnProvisionImage(object sender, ProvisionImageEventArgs e)
        {
            string filename = Path.GetFileName(e.ImageUrl.OriginalString);
            if (!File.Exists("../../images/" + filename))
            {
                e.Cancel = true;
                return;
            }

            e.Provision(File.ReadAllBytes("../../images/" + filename));
        }

        static void AssertThatOpenXmlDocumentIsValid(WordprocessingDocument wpDoc)
        {
            var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
            var errors = validator.Validate(wpDoc);

            if (!errors.GetEnumerator().MoveNext())
                return;

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The document doesn't look 100% compatible with Office 2010.\n");

            Console.ForegroundColor = ConsoleColor.Gray;
            foreach (ValidationErrorInfo error in errors)
            {
                Console.Write("{0}\n\t{1}", error.Path.XPath, error.Description);
                Console.WriteLine();
            }

            Console.ReadLine();
        }
    }
}