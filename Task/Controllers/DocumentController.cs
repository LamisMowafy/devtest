
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Task.Models;
using DocumentFormat.OpenXml.ExtendedProperties;


public class DocumentController : Controller
{
    // GET: Word/Upload
    public IActionResult Upload()
    {
        return View();
    }

    // POST: Word/Upload
    [HttpPost]
    public async Task<IActionResult> Upload(IFormFile[] files)
    {
        if (files == null || files.Length == 0)
        {
            return View();
        }

        var report = new Reports();

        foreach (var file in files)
        {
            if (file.Length > 0)
            {
                var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", file.FileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                var metadata = ExtractMetadata(filePath);
                var wordcount = CountWords(filePath);

                report.ProcessedFiles.Add(file.FileName);
                if (metadata.MissingMetadata)
                    report.FilesWithMissingMetadata.Add(file.FileName);

                report.TotalWordCount += wordcount;
            }
        }

        var reportJson = JsonConvert.SerializeObject(report, Formatting.Indented);
        var reportFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "reports", "word_Report.json");
        System.IO.File.WriteAllText(reportFilePath, reportJson);

        return File(System.IO.File.ReadAllBytes(reportFilePath), "application/json", "word_Report.json");
    }

    private MetadataModel ExtractMetadata(string filePath)
    {
        using (var wordDoc = WordprocessingDocument.Open(filePath, false))
        {
            var coreProps = wordDoc.ExtendedProperties.Properties;
            var title = coreProps.Elements<Title>().FirstOrDefault()?.Text;
            var author = coreProps.Elements<Author>().FirstOrDefault()?.Text;
            var creationDate = coreProps.Elements<CreationDate>().FirstOrDefault()?.Text;

            bool missingMetadata = string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(author) || string.IsNullOrWhiteSpace(creationDate);

            return new MetadataModel
            {
                Title = title,
                Author = author,
                CreationDate = creationDate,
                MissingMetadata = missingMetadata
            };
        }
    }
    private int CountWords(string filePath)
    {
        int wordCount = 0;
        using (var wordDoc = WordprocessingDocument.Open(filePath, false))
        {
            var body = wordDoc.MainDocumentPart.Document.Body;
            var text = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                           .Where(t => t.Text != null)
                           .Select(t => t.Text);
            wordCount = text.Sum(t => t.Split(new char[] { ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length);
        }
        return wordCount;
    }
}
