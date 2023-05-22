using System;

namespace BrandixAutomation.Labdip.API.Models
{
    public class HeaderMaster
    {
        public int Id { get; set; }
        public string HeaderType { get; set; }
        public int Variation { get; set; }
        public string FileName { get; set; }
        public string HeaderAttribute { get; set; }
        public string SubHeaderAttribute { get; set; }
        public Boolean SetSameHeaderasValue { get; set; }
        public string SubValueAttribute { get; set; }
        public string LowerBoundHedaerName { get; set; }
        public string UpperBoundHeaderName { get; set; }
        public string HedaerName { get; set; }
        public int ColumnSpan { get; set; }
        public int RowSpan { get; set; }
        public int SubId { get; set; }
        public int ColumnNo { get; set; }
        public int RowNo { get; set; }
        public int LowerBoundColumnNo { get; set; }
        public int LowerBoundRowNo { get; set; }
        public int UpperBoundColumnNo { get; set; }
        public int UpperBoundRowNo { get; set; }
        public Boolean UpdateforAll { get; set; }
        public string FilterData { get; set; }
        public Boolean SkipBlanks { get; set; }
        public Boolean RepeatData { get; set; }
        public string Split { get; set; }
        public string Replace { get; set; }
        public string Extract { get; set; }
        public string TransformData { get; set; }
        public Boolean SaveTransformVariation { get; set; }
    }
}
