using BrandixAutomation.Labdip.API.Models;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace BrandixAutomation.Labdip.API.ProcessFiles
{
    public class OptionService
    {
        List<Options> _optionsList;
        string _filePathName = "LabdipChartConfigJSONs/OptionsList.json";

        public OptionService()
        {
            _optionsList = ReadJson();

        }

        public List<Options> GetOptions()
        {
            return _optionsList;
        }

        public Options GetOptionById(string OptionId)
        {
            return _optionsList.Where(c => c.OptionId == OptionId).FirstOrDefault();
        }


        public bool InsertNewRecord(Options options)
        {
            if (options != null)
            {
                _optionsList.Add(options);
                WriteJson(_optionsList);
            }
            return true;
        }

        public bool UpdateRecord(Options options)
        {
            if (options != null)
            {
                _optionsList.ForEach(ele =>
                {
                    if (ele.OptionId == options.OptionId)
                    {
                        ele.OptionName = options.OptionName;
                    }
                });
                WriteJson(_optionsList);
            }
            return true;
        }

        public bool DeleteRecord(Options options)
        {
            if (options != null)
            {
                var index = _optionsList.FindIndex(c => c.OptionId == options.OptionId);
                _optionsList.RemoveAt(index);
                WriteJson(_optionsList);
            }
            return true;
        }

        private List<Options> ReadJson()
        {

            string jsonString = File.ReadAllText(_filePathName);
            List<Options> result = JsonSerializer.Deserialize<List<Options>>(jsonString);
            return result;
        }

        private bool WriteJson(List<Options> options)
        {
            var ORderedList = ReOrdeIds(options);
            string jsonString = JsonSerializer.Serialize<List<Options>>(options);
            File.WriteAllText(_filePathName, jsonString);
            return true;
        }

        private List<Options> ReOrdeIds(List<Options> options)
        {
            options.OrderBy(c => c.OptionId).ToList();

            return options;
        }
    }
}
