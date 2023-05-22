using System;

namespace BrandixAutomation.Labdip.API.DTOs
{
    public class Request
    {
        DateTime defaultTime = DateTime.UtcNow;
        public DateTime LastAccessDateTime
        {
            get { return this.defaultTime; }
            set { value = this.defaultTime; }
        }
    }
}
