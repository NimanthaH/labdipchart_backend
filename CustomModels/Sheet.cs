using System.Collections.Generic;

namespace BrandixAutomation.Labdip.API.Models
{
    public class Sheet
    {
        public List<Row> Rows { get; set; }
        public string SheetName { get; set; }

        public Sheet(string sheetName)
        {
            Rows = new List<Row>();
            SheetName = sheetName;
        }
    }
}
