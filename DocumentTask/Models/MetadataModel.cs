namespace DocumentTask.Models
{
    public class MetadataModel
    {
        public string FilePath { get; set; }
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime? CreationDate { get; set; }
        public int PageCount { get; set; }
        public int WordCount { get; set; }
    }
}
