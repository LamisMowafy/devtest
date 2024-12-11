namespace Task.Models
{
    public class Reports
    {
        public List<string> ProcessedFiles { get; set; } = new List<string>();
        public List<string> FilesWithMissingMetadata { get; set; } = new List<string>();
        public int TotalWordCount { get; set; } = 0;
    }
}
