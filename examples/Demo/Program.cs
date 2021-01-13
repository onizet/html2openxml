using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            const string filename = "test.docx";
            string html = ResourceHelper.GetString("Resources.CompleteRunTest.html");
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
                    //converter.WebProxy.Credentials = new System.Net.NetworkCredential("nizeto", "****", "domain");
                    //converter.WebProxy.Proxy = new System.Net.WebProxy("proxy01:8080");
                    converter.ImageProcessing = ImageProcessing.ManualProvisioning;
                    converter.ProvisionImage += OnProvisionImage;
                    converter.BeforeProcess +=  OnBeforeProcess;
                    converter.AfterProcess += Converter_AfterProcess;
                    Body body = mainPart.Document.Body;

                    converter.ParseHtml(html);
                    mainPart.Document.Save();

                    AssertThatOpenXmlDocumentIsValid(package);
                }

                File.WriteAllBytes(filename, generatedDocument.ToArray());
            }

            System.Diagnostics.Process.Start(filename);
        }

        private static void Converter_AfterProcess(object sender, AfterProcessEventArgs e)
        {
            switch (e.Tag)
            {
                case "<img>":
                    if (e.CurrentParagraph == null)
                        return;
                    if (e.HtmlAttributes["class"] == null)
                        return;
                    ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                    if (e.HtmlAttributes["class"].Contains("fr-fic"))
                        paragraphProperties1.Append(new Justification() { Val = JustificationValues.Center });
                    if (e.HtmlAttributes["class"].Contains("fr-fil"))
                        paragraphProperties1.Append(new Justification() { Val = JustificationValues.Left });
                    if (e.HtmlAttributes["class"].Contains("fr-fir"))
                        paragraphProperties1.Append(new Justification() { Val = JustificationValues.Right });
                    e.CurrentParagraph.Append(paragraphProperties1);
                    break;
            }
        }

        /// <summary>
        /// Example on how to change attributes of an html Tag if necessary in this case all spans getting bold
        /// </summary>
        /// <param name="current"></param>
        /// <param name="Tag"></param>
        static void OnBeforeProcess(object sender, BeforeProcessEventArgs e)
        {
            switch (e.Tag)
            {
                case "<span>":
                    if (!string.IsNullOrEmpty(e.Current["style"]))
                        e.Current["style"] = "font-weight: bold";
                    break;
            }
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