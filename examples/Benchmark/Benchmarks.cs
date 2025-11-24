using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

[MemoryDiagnoser]
//[SimpleJob(runtimeMoniker: RuntimeMoniker.Net48)]
[SimpleJob(runtimeMoniker: RuntimeMoniker.Net80, baseline: true)]
public class Benchmarks
{
    [Benchmark]
    public async Task ParseWithSpan()
    {
        string html = ResourceHelper.GetString("benchmark.html");

        using (MemoryStream generatedDocument = new MemoryStream())
        using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
        {
            MainDocumentPart? mainPart = package.MainDocumentPart;
            if (mainPart == null)
            {
                mainPart = package.AddMainDocumentPart();
                new Document(new Body()).Save(mainPart);
            }

            HtmlConverter converter = new HtmlConverter(mainPart);
            converter.RenderPreAsTable = true;

            await converter.ParseBody(html);
            mainPart.Document.Save();
        }
    }
}