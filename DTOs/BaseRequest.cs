namespace BrandixAutomation.Labdip.API.DTOs
{
    public class BaseRequest<T> : Request
    {
        public T Request { get; set; }
    }
}
