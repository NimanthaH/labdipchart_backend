using BrandixAutomation.Labdip.API.DTOs;
using BrandixAutomation.Labdip.API.Models;
using BrandixAutomation.Labdip.API.ProcessFiles;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;

namespace BrandixAutomation.Labdip.API.Controllers
{
    //dotnet run
    [Route("api/[controller]")]
    [ApiController]
    public class LabdipController : ControllerBase
    {
        [HttpGet]
        public IActionResult Get()
        {
            //Test Api is Up and Running
            return StatusCode(200, "Labdip Chart Api Connected!");
        }

        #region Labdip

        [HttpPost("labdipChart"), DisableRequestSizeLimit]
        public IActionResult LabdipChartProcess()
        {
            IActionResult result = Ok();
            LabdipChartDataService service = new LabdipChartDataService();

            try
            {
                var files = Request.Form.Files;
                var keys = Request.Form.Keys;
                if (files != null && files.Count > 0)
                {
                    return Ok(service.GetLabdipChartData(files, keys));
                }
                else
                {
                    return BadRequest();
                }
            }
            catch (System.Exception ex)
            {
                return StatusCode(500, $"Internal Server error:{ex.Message}");
            }
        }


        [HttpGet("options")]
        public AutomationResponse<List<Options>> GetOptions()
        {
            var response = new AutomationResponse<List<Options>>();
            try
            {
                OptionService optionService = new OptionService();
                response.Data = optionService.GetOptions();
            }
            catch (Exception ex)
            {

                response.SetResponseStatus(false, ex);
            }
            return response;
        }

        [HttpGet("customers")]
        public AutomationResponse<List<Customers>> GetCustomers()
        {
            var response = new AutomationResponse<List<Customers>>();
            try
            {
                CustomerService customerService = new CustomerService();
                response.Data = customerService.GetCustomers();
            }
            catch (Exception ex)
            {

                response.SetResponseStatus(false, ex);
            }
            return response;
        }

        [HttpGet("headermaster")]
        public AutomationResponse<List<HeaderMaster>> GetHeaderMaster()
        {
            var response = new AutomationResponse<List<HeaderMaster>>();
            try
            {
                HeaderMasterService headerMasterService = new HeaderMasterService();
                response.Data = headerMasterService.GetHeaderList();
            }
            catch (Exception ex)
            {

                response.SetResponseStatus(false, ex);
            }
            return response;
        }

        #endregion


        #region Tread

        [HttpPost("threadShade"), DisableRequestSizeLimit]
        public AutomationResponse<ThreadShadeResponse> ThreadShadeProcess()
        {
            var response = new AutomationResponse<ThreadShadeResponse>();
            try
            {
                if (Request.Form.Files.Count > 1)
                {
                    var labdipChart = Request.Form.Files[0];
                    var threadShade = Request.Form.Files[1];
                    string threadTypes = Convert.ToString(Request.Form["ThreadTypes"]);

                    ThreadShadeDataService threadShadeDataService = new ThreadShadeDataService(labdipChart, threadShade, threadTypes);
                    response.Data = threadShadeDataService.ProcessThreadShadeData();
                }
                else
                    throw new Exception("Bad Request or File Error");
            }
            catch (Exception ex)
            {
                response.SetResponseStatus(false, ex);
            }
            return response;
        }

        [HttpGet("threadTypes")]
        public AutomationResponse<List<ThreadTypes>> GetThreadTypes()
        {
            var response = new AutomationResponse<List<ThreadTypes>>();
            try
            {
                ThreadTypeService threadTypeService = new ThreadTypeService();
                response.Data = threadTypeService.GetThreadTypes();
            }
            catch (Exception ex)
            {

                response.SetResponseStatus(false, ex);
            }
            return response;
        }

        [HttpPost("insertNewThread")]
        public AutomationResponse<List<ThreadTypes>> InsertNewThread([FromBody] AutomationRequest<ThreadTypes> request)
        {
            var response = new AutomationResponse<List<ThreadTypes>>();
            try
            {
                ThreadTypeService threadTypeService = new ThreadTypeService();
                if (threadTypeService.InsertNewRecord(request.Request))
                {
                    response.Data = threadTypeService.GetThreadTypes();
                }
            }
            catch (Exception ex)
            {
                response.SetResponseStatus(false, ex);
            }
            return response;
        }

        [HttpPost("updateThread")]
        public AutomationResponse<List<ThreadTypes>> UpdateThread([FromBody] AutomationRequest<ThreadTypes> request)
        {
            var response = new AutomationResponse<List<ThreadTypes>>();
            try
            {
                ThreadTypeService threadTypeService = new ThreadTypeService();
                if (threadTypeService.UpdateRecord(request.Request))
                {
                    response.Data = threadTypeService.GetThreadTypes();
                }
            }
            catch (Exception ex)
            {
                response.SetResponseStatus(false, ex);
            }
            return response;
        }

        [HttpPost("deleteThread")]
        public AutomationResponse<List<ThreadTypes>> deleteThread([FromBody] AutomationRequest<ThreadTypes> request)
        {
            var response = new AutomationResponse<List<ThreadTypes>>();
            try
            {
                ThreadTypeService threadTypeService = new ThreadTypeService();
                if (threadTypeService.DeleteRecord(request.Request))
                {
                    response.Data = threadTypeService.GetThreadTypes();
                }
            }
            catch (Exception ex)
            {
                response.SetResponseStatus(false, ex);
            }
            return response;
        }

        #endregion

    }
}
