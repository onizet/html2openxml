﻿using System;
using System.IO;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.Threading.Tasks;

namespace Demo
{
    static class Program
    {
        static async Task Main(string[] args)
        {
            const string filename = "test.docx";
            string html = ResourceHelper.GetString("Resources.Document.html");
            if (File.Exists(filename)) File.Delete(filename);

            using (MemoryStream generatedDocument = new MemoryStream())
            {
                // Uncomment and comment the second using() to open an existing template document
                // instead of creating it from scratch.
                using (var buffer = ResourceHelper.GetStream("Resources.template.docx"))
                {
                    buffer.CopyTo(generatedDocument);
                }

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
                   // HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
                    //FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
                    converter.RenderPreAsTable = true;
                    Body body = mainPart.Document.Body;

                    await converter.ParseHeader(@"<a href=""www.github.com"">
                        <img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg=="" alt=""Red dot"" /> Red Dot</a>");

                    await converter.ParseBody(html);
                    mainPart.Document.Save();

                    AssertThatOpenXmlDocumentIsValid(package);
                }

                File.WriteAllBytes(filename, generatedDocument.ToArray());
            }

            Process.Start(new ProcessStartInfo(filename) { UseShellExecute = true });
        }

        static void AssertThatOpenXmlDocumentIsValid(WordprocessingDocument wpDoc)
        {
            var validator = new OpenXmlValidator(FileFormatVersions.Office2021);
            var errors = validator.Validate(wpDoc);

            if (!errors.GetEnumerator().MoveNext())
                return;

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The document doesn't look 100% compatible with Office 2021.\n");

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