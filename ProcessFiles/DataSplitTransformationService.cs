using BrandixAutomation.Labdip.API.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace BrandixAutomation.Labdip.API.ProcessFiles
{
    public class DataSplitTransformationService
    {
        List<DataSplitTransformation> _dataSplitTransformationList;
        string _filePathName = "LabdipChartConfigJSONs/DataSplitTransformationList.json";

        public DataSplitTransformationService()
        {
            _dataSplitTransformationList = ReadJson();
        }

        public List<DataSplitTransformation> GetDataSplitTransformationList()
        {
            return _dataSplitTransformationList;
        }

        public List<DataSplitTransformation> GetDataSplitTransformationListbyValue(string value, int variation)
        {
            return _dataSplitTransformationList.Where(c => c.InitialData == value && c.Variation == variation).ToList();
        }

        public int GetMaxTranformationId()
        {
            return _dataSplitTransformationList.Max(c => c.Id) + 1;
        }


        public bool InsertNewRecord(DataSplitTransformation transformation)
        {
            try
            {

                if (transformation != null)
                {
                    _dataSplitTransformationList.Add(transformation);
                    WriteJson(_dataSplitTransformationList);
                }
                return true;
            }
            catch(Exception error)
            {
                return false;
            }
        }

        public bool UpdateRecord(DataSplitTransformation transformation)
        {
            if (transformation != null)
            {
                _dataSplitTransformationList.ForEach(ele =>
                {
                    if (ele.Id == transformation.Id)
                    {
                        ele.SubId = transformation.SubId;
                        ele.Variation = transformation.Variation;
                        ele.HeaderAttribute = transformation.HeaderAttribute;
                        ele.InitialData = transformation.InitialData;
                        ele.TransformedData = transformation.TransformedData;
                    }
                });
                WriteJson(_dataSplitTransformationList);
            }
            return true;
        }

        public DataSplitTransformation UpdateModelObject(DataSplitTransformation transformation)
        {
            DataSplitTransformation transformation_output = new DataSplitTransformation();
            if (transformation != null)
            {
                transformation_output.Id = transformation.Id;
                transformation_output.SubId = transformation.SubId;
                transformation_output.Variation = transformation.Variation;
                transformation_output.HeaderAttribute = transformation.HeaderAttribute;
                transformation_output.InitialData = transformation.InitialData;
                transformation_output.TransformedData = transformation.TransformedData;
            }

            return transformation_output;
        }

        public bool DeleteRecord(DataSplitTransformation transformation)
        {
            if (transformation != null)
            {
                var index = _dataSplitTransformationList.FindIndex(c => c.Id == transformation.Id && c.Variation == transformation.Variation);
                _dataSplitTransformationList.RemoveAt(index);
                WriteJson(_dataSplitTransformationList);
            }
            return true;
        }

        private List<DataSplitTransformation> ReadJson()
        {

            string jsonString = File.ReadAllText(_filePathName);
            List<DataSplitTransformation> result = JsonSerializer.Deserialize<List<DataSplitTransformation>>(jsonString);
            return result;
        }

        private bool WriteJson(List<DataSplitTransformation> transformations)
        {
            var ORderedList = ReOrdeIds(transformations);
            string jsonString = JsonSerializer.Serialize<List<DataSplitTransformation>>(transformations);
            File.WriteAllText(_filePathName, jsonString);
            return true;
        }

        private List<DataSplitTransformation> ReOrdeIds(List<DataSplitTransformation> transformations)
        {
            for (int r = 0; r < transformations.Count; r++)
            {
                transformations[r].Id = r;
            }
            return transformations;
        }
    }
}
