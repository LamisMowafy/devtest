using DocumentFormat.OpenXml.Packaging;
using DocumentTask.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace DocumentTask.Controllers
{
    public class DocumentController : Controller
    {

        // View for uploading Word files
        public IActionResult Index()
        {

            return View();
        }

        // Handle file upload, process files, and generate the report
        [HttpPost]
        public async Task<IActionResult> ProcessFiles(List<IFormFile> files)
        {
            if (files == null || files.Count == 0)
            {
                ModelState.AddModelError("", "No files selected.");
                return View("Index");
            }

            var wordReport = await ProcessDocuments(files);
            var jsonReport = JsonConvert.SerializeObject(wordReport, Newtonsoft.Json.Formatting.Indented);

            // Save the JSON report
            var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "reports", "word_Report.json");
            System.IO.File.WriteAllText(reportPath, jsonReport);
            return File(System.IO.File.ReadAllBytes(reportPath), "application/json", "word_Report.json");
        }

        // Process a list of Word documents and extract metadata
        private async Task<Reports> ProcessDocuments(List<IFormFile> files)
        {
            var report = new Reports();
            var tasks = files.Select(file => ProcessWordDocument(file, report)).ToList();
            await Task.WhenAll(tasks);
            return report;
        }

        // Extract metadata and calculate word count from each document
        private async Task ProcessWordDocument(IFormFile file, Reports report)
        {
            try
            {
                using (var stream = file.OpenReadStream())
                {
                    var metadata = ExtractMetadata(stream);
                    metadata.FilePath = file.FileName;
                    report.ProcessedFiles.Add(file.FileName);

                    // Check for missing metadata
                    if (string.IsNullOrEmpty(metadata.Title) || string.IsNullOrEmpty(metadata.Author) || !metadata.CreationDate.HasValue)
                    {
                        report.FilesWithMissingMetadata.Add(file.FileName);
                    }

                    // Calculate total word count
                    metadata.WordCount = await GetWordCount(stream);

                    // Add word count to total
                    report.TotalWordCount += metadata.WordCount;
                }
            }
            catch (Exception ex)
            {
                report.Errors.Add($"Error processing {file.FileName}: {ex.Message}");
            }
        }

        // Extract metadata from the Word document
        private MetadataModel ExtractMetadata(Stream stream)
        {
            var metadata = new MetadataModel();

            using (var doc = WordprocessingDocument.Open(stream, false))
            {
                var coreProperties = doc.PackageProperties;

                metadata.Title = coreProperties.Title;
                metadata.Author = coreProperties.Creator;
                metadata.CreationDate = coreProperties.Created;
            }

            return metadata;
        }

        // Count words in the Word document
        private async Task<int> GetWordCount(Stream stream)
        {
            int wordCount = 0;

            using (var doc = WordprocessingDocument.Open(stream, false))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var textElements = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>();

                wordCount = textElements.Where(t => !string.IsNullOrEmpty(t.Text))
                                        .Sum(t => t.Text.Split(new[] { ' ', '\t', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length);
            }

            return wordCount;
        }


    }
}
