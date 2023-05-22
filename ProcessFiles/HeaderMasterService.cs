using BrandixAutomation.Labdip.API.Models;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace BrandixAutomation.Labdip.API.ProcessFiles
{
    public class HeaderMasterService
    {
        List<HeaderMaster> _headerMasterList;
        string _filePathName = "LabdipChartConfigJSONs/HeaderMaster.json";

        public HeaderMasterService()
        {
            _headerMasterList = ReadJson();
        }

        public List<HeaderMaster> GetHeaderList()
        {
            return _headerMasterList;
        }


        public bool InsertNewRecord(HeaderMaster head)
        {
            if (head != null)
            {
                _headerMasterList.Add(head);
                WriteJson(_headerMasterList);
            }
            return true;
        }

        public bool UpdateRecord(HeaderMaster head)
        {
            if (head != null)
            {
                _headerMasterList.ForEach(ele =>
                {
                    if (ele.Id == head.Id)
                    {
                        ele.HeaderType = head.HeaderType;
                        ele.Variation = head.Variation;
                        ele.FileName = head.FileName;
                        ele.HeaderAttribute = head.HeaderAttribute;
                        ele.SubHeaderAttribute = head.SubHeaderAttribute;
                        ele.SubValueAttribute = head.SubValueAttribute;
                        ele.SetSameHeaderasValue = head.SetSameHeaderasValue;
                        ele.LowerBoundHedaerName = head.LowerBoundHedaerName;
                        ele.UpperBoundHeaderName = head.UpperBoundHeaderName;
                        ele.HedaerName = head.HedaerName;
                        ele.ColumnSpan = head.ColumnSpan;
                        ele.RowSpan = head.RowSpan;
                        ele.SubId = head.SubId;
                        ele.ColumnNo = head.ColumnNo;
                        ele.RowNo = head.RowNo;
                        ele.UpperBoundColumnNo = head.UpperBoundColumnNo;
                        ele.UpperBoundRowNo = head.UpperBoundRowNo;
                        ele.LowerBoundColumnNo = head.LowerBoundColumnNo;
                        ele.LowerBoundRowNo = head.LowerBoundRowNo;
                        ele.UpdateforAll = head.UpdateforAll;
                        ele.FilterData = head.FilterData;
                        ele.RepeatData = head.RepeatData;
                        ele.SkipBlanks = head.SkipBlanks;
                        ele.Split = head.Split;
                        ele.Replace = head.Replace;
                        ele.Extract = head.Extract;
                        ele.TransformData = head.TransformData;
                        ele.SaveTransformVariation = head.SaveTransformVariation;
                    }
                });
                WriteJson(_headerMasterList);
            }
            return true;
        }

        public HeaderMaster UpdateModelObject(HeaderMaster head)
        {
            HeaderMaster head_output = new HeaderMaster();
            if (head != null)
            {
                head_output.Id = head.Id;
                head_output.HeaderType = head.HeaderType;
                head_output.Variation = head.Variation;
                head_output.FileName = head.FileName;
                head_output.HeaderAttribute = head.HeaderAttribute;
                head_output.SubHeaderAttribute = head.SubHeaderAttribute;
                head_output.SubValueAttribute = head.SubValueAttribute;
                head_output.SetSameHeaderasValue = head.SetSameHeaderasValue;
                head_output.LowerBoundHedaerName = head.LowerBoundHedaerName;
                head_output.UpperBoundHeaderName = head.UpperBoundHeaderName;
                head_output.HedaerName = head.HedaerName;
                head_output.ColumnSpan = head.ColumnSpan;
                head_output.RowSpan = head.RowSpan;
                head_output.SubId = head.SubId;
                head_output.ColumnNo = head.ColumnNo;
                head_output.RowNo = head.RowNo;
                head_output.UpperBoundColumnNo = head.UpperBoundColumnNo;
                head_output.UpperBoundRowNo = head.UpperBoundRowNo;
                head_output.LowerBoundColumnNo = head.LowerBoundColumnNo;
                head_output.LowerBoundRowNo = head.LowerBoundRowNo;
                head_output.UpdateforAll = head.UpdateforAll;
                head_output.FilterData = head.FilterData;
                head_output.RepeatData = head.RepeatData;
                head_output.SkipBlanks = head.SkipBlanks;
                head_output.Split = head.Split;
                head_output.Replace = head.Replace;
                head_output.Extract = head.Extract;
                head_output.TransformData = head.TransformData;
                head_output.SaveTransformVariation = head.SaveTransformVariation;
            }

            return head_output;
        }

        public bool DeleteRecord(HeaderMaster head)
        {
            if (head != null)
            {
                var index = _headerMasterList.FindIndex(c => c.Id == head.Id && c.SubId == head.SubId);
                _headerMasterList.RemoveAt(index);
                WriteJson(_headerMasterList);
            }
            return true;
        }

        private List<HeaderMaster> ReadJson()
        {

            string jsonString = File.ReadAllText(_filePathName);
            List<HeaderMaster> result = JsonSerializer.Deserialize<List<HeaderMaster>>(jsonString);
            return result;
        }

        private bool WriteJson(List<HeaderMaster> headerMaster)
        {
            var ORderedList = ReOrdeIds(headerMaster);
            string jsonString = JsonSerializer.Serialize<List<HeaderMaster>>(headerMaster);
            File.WriteAllText(_filePathName, jsonString);
            return true;
        }

        private List<HeaderMaster> ReOrdeIds(List<HeaderMaster> headers)
        {
            for (int r = 0; r < headers.Count; r++)
            {
                headers[r].Id = r;
            }
            return headers;
        }
    }
}
