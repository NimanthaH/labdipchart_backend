using BrandixAutomation.Labdip.API.Models;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace BrandixAutomation.Labdip.API.ProcessFiles
{
    public class CustomerService
    {
        List<Customers> _customerList;
        string _filePathName = "LabdipChartConfigJSONs/CustomerList.json";

        public CustomerService()
        {
            _customerList = ReadJson();
        }

        public List<Customers> GetCustomers()
        {
            return _customerList;
        }


        public bool InsertNewRecord(Customers customer)
        {
            if (customer != null)
            {
                _customerList.Add(customer);
                WriteJson(_customerList);
            }
            return true;
        }

        public bool UpdateRecord(Customers customer)
        {
            if (customer != null)
            {
                _customerList.ForEach(ele =>
                {
                    if (ele.Id == customer.Id)
                    {
                        ele.Name = customer.Name;
                        ele.Variation = customer.Variation;
                    }
                });
                WriteJson(_customerList);
            }
            return true;
        }

        public bool DeleteRecord(Customers customer)
        {
            if (customer != null)
            {
                var index = _customerList.FindIndex(c => c.Id == customer.Id);
                _customerList.RemoveAt(index);
                WriteJson(_customerList);
            }
            return true;
        }

        private List<Customers> ReadJson()
        {

            string jsonString = File.ReadAllText(_filePathName);
            List<Customers> result = JsonSerializer.Deserialize<List<Customers>>(jsonString);
            return result;
        }

        private bool WriteJson(List<Customers> customers)
        {
            var ORderedList = ReOrdeIds(customers);
            string jsonString = JsonSerializer.Serialize<List<Customers>>(customers);
            File.WriteAllText(_filePathName, jsonString);
            return true;
        }

        private List<Customers> ReOrdeIds(List<Customers> placements)
        {
            for (int r = 0; r < placements.Count; r++)
            {
                placements[r].Id = r;
            }
            return placements;
        }
    }
}
