using Microsoft.AspNetCore.Http;
using System.ComponentModel.DataAnnotations;

namespace BrandixAutomation.Labdip.API.DTOs
{
    public class ThreadShadeRequest
    {
        [Required]
        public IFormFileCollection File { get; set; }
        [Required]
        public string[] ThreadTypes { get; set; }
    }


}
