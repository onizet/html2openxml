using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;

[MemoryDiagnoser]
//[SimpleJob(runtimeMoniker: RuntimeMoniker.Net462)]
//[SimpleJob(runtimeMoniker: RuntimeMoniker.Net50)]
[SimpleJob(runtimeMoniker: RuntimeMoniker.Net80)]
public class Benchmarks
{
    [Benchmark]
    public async Task ParseWithSpan()
    {
        string html = ResourceHelper.GetString("benchmark.html");

        using (MemoryStream generatedDocument = new MemoryStream())
        using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = package.MainDocumentPart;
            if (mainPart == null)
            {
                mainPart = package.AddMainDocumentPart();
                new Document(new Body()).Save(mainPart);
            }

            HtmlConverter converter = new HtmlConverter(mainPart);
            converter.RenderPreAsTable = true;
            Body body = mainPart.Document.Body!;

            await converter.ParseBody(html);
            mainPart.Document.Save();
        }
    }
}