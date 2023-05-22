using Microsoft.Extensions.Logging;
using System;
using System.Net;

namespace BrandixAutomation.Labdip.API.DTOs
{
    public class AutomationResponse<T> : BaseResponse<T>
    {
        private ILogger _iLogger;

        public AutomationResponse(ILogger logger = null, bool? isStatus = null) : base(isStatus)
        {
            _iLogger = logger;
        }

        public override void SetResponseStatus(bool isSuccess = true, Exception ex = null)
        {
            if ((isSuccess) && (ex == null)) IsSuccess = isSuccess;
            else
            {
                IsSuccess = false;
                Message = ex.Message;
                MessageDetails = ex.StackTrace;
                StatusCode = (int)HttpStatusCode.InternalServerError;
                SendToLog();
            }
        }

        private void SendToLog()
        {
            //call Logging Repository ..SeriLog or Kibana..
        }



    }

}
