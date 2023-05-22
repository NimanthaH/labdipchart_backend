using System.Collections.Generic;

namespace BrandixAutomation.Labdip.API.Models
{
    public class LabdipChartModel
    {
        public int IDX { get; set; }
        public int Index { get; set; }
        public int RowIndex { get; set; }
        public string Division { get; set; }
        public string Season { get; set; }
        public string Category { get; set; }
        public string Program { get; set; }
        public string StyleNoIndividual { get; set; }
        public string GMTDescription { get; set; }
        public string GMTColor { get; set; }
        public string NRF { get; set; }
        public string ColorCode { get; set; }
        public string RMColor { get; set; }
        public string PackCombination { get; set; }
        public string PalcementName { get; set; }
        public string BOMSelection { get; set; }
        public string ItemName { get; set; }
        public string SupplierName { get; set; }
        public string RMColorRef { get; set; }
        public string GarmentWay { get; set; }
        public string FBNumber { get; set; }
        public string MaterialType { get; set; }
        public string ColorDyeingTechnic { get; set; }
        public List<LabdipChartSubModel> SubColumns { get; set; }
        public string error { get; set; }
        public string Split { get; set; }
    }
}
