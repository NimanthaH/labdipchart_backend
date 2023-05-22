using Microsoft.AspNetCore.Http;
using System.ComponentModel.DataAnnotations;

namespace BrandixAutomation.Labdip.API.DTOs
{
    public class LabdipProcessRequest
    {
        [Required]
        public IFormCollection File { get; set; }

    }
}
