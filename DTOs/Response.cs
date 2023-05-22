using System;
using System.Net;

namespace BrandixAutomation.Labdip.API.DTOs
{
    public class Response
    {
        DateTime defaultTime = DateTime.UtcNow;

        public string Message { get; set; }
        public string MessageDetails { get; set; }
        public ResponseStatus Status { get; set; }
        public DateTime LastAccessedDateTime
        {
            get { return this.defaultTime; }
            set { value = this.defaultTime; }
        }
        public bool IsSuccess { get; set; }

        public int StatusCode { get; set; }
        public virtual void SetResponseStatus(bool isSuccess = true, Exception ex = null)
        {
            if ((isSuccess) && (ex == null)) IsSuccess = isSuccess;
            else
            {
                IsSuccess = false;
                Message = ex.Message;
                MessageDetails = ex.StackTrace;
                Status = ResponseStatus.Error;
                StatusCode = (int)HttpStatusCode.InternalServerError;
            }
        }
    }
}
