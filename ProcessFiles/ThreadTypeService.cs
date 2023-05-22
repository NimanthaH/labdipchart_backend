using BrandixAutomation.Labdip.API.Models;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace BrandixAutomation.Labdip.API.ProcessFiles
{
    public class ThreadTypeService
    {
        List<ThreadTypes> _threadTypeList;
        string _filePathName = "LabdipChartConfigJSONs/ThreadTypesList.json";

        public ThreadTypeService()
        {
            _threadTypeList = ReadJson();

        }

        public List<ThreadTypes> GetThreadTypes()
        {
            return _threadTypeList;
        }


        public bool InsertNewRecord(ThreadTypes thread)
        {
            if (thread != null)
            {
                _threadTypeList.Add(thread);
                WriteJson(_threadTypeList);
            }
            return true;
        }

        public bool UpdateRecord(ThreadTypes thread)
        {
            if (thread != null)
            {
                _threadTypeList.ForEach(ele =>
                {
                    if (ele.Id == thread.Id)
                    {
                        ele.ThreadType = thread.ThreadType;
                        ele.Supplier = thread.Supplier;
                        ele.TicketNo = thread.TicketNo;
                    }
                });
                WriteJson(_threadTypeList);
            }
            return true;
        }

        public bool DeleteRecord(ThreadTypes thread)
        {
            if (thread != null)
            {
                var index = _threadTypeList.FindIndex(c => c.ThreadType == thread.ThreadType);
                _threadTypeList.RemoveAt(index);
                WriteJson(_threadTypeList);
            }
            return true;
        }

        private List<ThreadTypes> ReadJson()
        {

            string jsonString = File.ReadAllText(_filePathName);
            List<ThreadTypes> result = JsonSerializer.Deserialize<List<ThreadTypes>>(jsonString);
            return result;
        }

        private bool WriteJson(List<ThreadTypes> threadTypes)
        {
            var ORderedList = ReOrdeIds(threadTypes);
            string jsonString = JsonSerializer.Serialize<List<ThreadTypes>>(threadTypes);
            File.WriteAllText(_filePathName, jsonString);
            return true;
        }

        private List<ThreadTypes> ReOrdeIds(List<ThreadTypes> threads)
        {
            for (int r = 0; r < threads.Count; r++)
            {
                threads[r].Id = r;
            }
            return threads;
        }
    }
}
