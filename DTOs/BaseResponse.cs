using System.Net;

namespace BrandixAutomation.Labdip.API.DTOs
{
    public class BaseResponse<T> : Response
    {
        public BaseResponse(bool? isStatus = null)
        {
            IsSuccess = isStatus ?? true;
            StatusCode = (int)HttpStatusCode.OK;
            Status = ResponseStatus.OK;
        }
        public T Data { get; set; }
    }
}
