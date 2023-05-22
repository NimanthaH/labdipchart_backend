using BrandixAutomation.Labdip.API.Models;
using System.Collections.Generic;

namespace BrandixAutomation.Labdip.API.DTOs
{
    public class ThreadShadeResponse
    {
        public IEnumerable<LabdipChartModel> LabdipChartModels { get; set; }
        public IEnumerable<LabdipChartModel> ProcessResult { get; set; }
        public IEnumerable<ThreadShadeModel> ThreadShadeModels { get; set; }
    }
}
