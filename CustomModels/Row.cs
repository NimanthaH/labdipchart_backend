using System.Collections.Generic;

namespace BrandixAutomation.Labdip.API.Models
{
    public class Row
    {
        public List<Cell> Cells { get; set; }
        public Row()
        {
            Cells = new List<Cell>();
        }
    }
}
