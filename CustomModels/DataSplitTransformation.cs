namespace BrandixAutomation.Labdip.API.Models
{
    public class DataSplitTransformation
    {
        public int Id { get; set; }
        public int SubId { get; set; }
        public int Variation { get; set; }
        public string HeaderAttribute { get; set; }
        public string InitialData { get; set; }
        public string TransformedData { get; set; }
    }
}
