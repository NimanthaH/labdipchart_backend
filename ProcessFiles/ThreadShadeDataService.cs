
using BrandixAutomation.Labdip.API.DTOs;
using BrandixAutomation.Labdip.API.Models;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;

namespace BrandixAutomation.Labdip.API.ProcessFiles
{
    public class ThreadShadeDataService
    {
        private Sheet _labdipChartSheet;
        private Sheet _threadShadeSheet;
        private List<string> _threadTypes;
        private List<ThreadTypes> _thredTypeList;

        public ThreadShadeDataService(IFormFile labdipChart, IFormFile threadShade, string threadTypes)
        {
            _labdipChartSheet = SetLabdipChart(labdipChart);
            _threadShadeSheet = SetThreadShade(threadShade);
            _threadTypes = SetThreadTypes(threadTypes);
            _thredTypeList = SetThredTypeList();
        }

        private Sheet SetLabdipChart(IFormFile file)
        {
            return ReadDataIntoSheet(file);
        }

        private Sheet SetThreadShade(IFormFile file)
        {
            return ReadDataIntoSheet(file);
        }

        private List<string> SetThreadTypes(string threadTypeString)
        {
            List<string> types = new List<string>();
            types.AddRange(threadTypeString.Split(","));
            return types;
        }

        private List<ThreadTypes> SetThredTypeList()
        {
            ThreadTypeService thredTypeService = new ThreadTypeService();
            return thredTypeService.GetThreadTypes();

        }

        private Sheet ReadDataIntoSheet(IFormFile file) //byte[] excelSheetByteArray
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var sheet = new Sheet("TempMaster");

            //var tempByteArray = FileToByteArray(@"C:\WorkSpaces\TechPacks.xlsx");           

            //using (var stream = new MemoryStream(tempByteArray))
            //{
            using (var reader = ExcelReaderFactory.CreateReader(file.OpenReadStream()))
            {
                //columnCount = reader.FieldCount;
                //rowCount = reader.RowCount;
                do
                {
                    int r = 0;
                    while (reader.Read())
                    {
                        var row = new Row();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row.Cells.Add(PrepareCell(i, reader.GetValue(i), ""));
                        }
                        r++;
                        sheet.Rows.Add(row);
                    }
                    reader.Close(); //only one Sheet reading..
                } while (reader.NextResult());
            }
            //}
            return sheet;
        }

        private Cell PrepareCell(int index, object cellValue, string NullReplacement)
        {
            return new Cell()
            {
                ColIndex = index,
                ColValue = cellValue != null ? cellValue.ToString() : NullReplacement
            };
        }

        #region Public Methods
        public ThreadShadeResponse ProcessThreadShadeData()
        {
            var labdip = SetLabdipChartData(_labdipChartSheet);
            var processedLabdip = SetLabdipChartData(_labdipChartSheet);
            var threadShade = SetThreadShadeChartData(_threadShadeSheet);

            return new ThreadShadeResponse()
            {
                LabdipChartModels = labdip,
                ThreadShadeModels = threadShade,
                ProcessResult = SetLabdipThreadShadeProcessData(processedLabdip, threadShade, _threadTypes)
            };

            //Result, Labdip, ThreadShade Objects as a single jason           
        }

        private List<LabdipChartModel> SetLabdipChartData(Sheet labdipDataSheet)
        {
            List<LabdipChartModel> result = new List<LabdipChartModel>();

            for (int i = 1; labdipDataSheet.Rows.Count > i; i++)
            {
                if (CheckRowDataVailabile(labdipDataSheet.Rows[i]))
                {
                    LabdipChartModel model = new LabdipChartModel()
                    {
                        Index = StringToInt(labdipDataSheet.Rows[i].Cells[1].ColValue),
                        Division = labdipDataSheet.Rows[i].Cells[2].ColValue,
                        Season = labdipDataSheet.Rows[i].Cells[3].ColValue,
                        Category = labdipDataSheet.Rows[i].Cells[4].ColValue,
                        Program = labdipDataSheet.Rows[i].Cells[5].ColValue,
                        StyleNoIndividual = labdipDataSheet.Rows[i].Cells[6].ColValue,
                        GMTDescription = labdipDataSheet.Rows[i].Cells[7].ColValue,
                        GMTColor = labdipDataSheet.Rows[i].Cells[8].ColValue,
                        NRF = labdipDataSheet.Rows[i].Cells[9].ColValue,
                        ColorCode = labdipDataSheet.Rows[i].Cells[10].ColValue,
                        PackCombination = labdipDataSheet.Rows[i].Cells[11].ColValue,
                        PalcementName = labdipDataSheet.Rows[i].Cells[12].ColValue,
                        BOMSelection = labdipDataSheet.Rows[i].Cells[13].ColValue,
                        ItemName = labdipDataSheet.Rows[i].Cells[14].ColValue,
                        SupplierName = labdipDataSheet.Rows[i].Cells[15].ColValue,
                        RMColor = labdipDataSheet.Rows[i].Cells[16].ColValue,
                        ColorDyeingTechnic = labdipDataSheet.Rows[i].Cells[17].ColValue,
                        RMColorRef = labdipDataSheet.Rows[i].Cells[18].ColValue,
                        GarmentWay = labdipDataSheet.Rows[i].Cells[19].ColValue,
                        FBNumber = labdipDataSheet.Rows[i].Cells[20].ColValue,
                        MaterialType = labdipDataSheet.Rows[i].Cells[21].ColValue
                    };
                    result.Add(model);
                }
            }
            return result;
        }

        private List<ThreadShadeModel> SetThreadShadeChartData(Sheet ThreadSahdeDataSheet)
        {
            List<ThreadShadeModel> result = new List<ThreadShadeModel>();

            for (int i = 1; ThreadSahdeDataSheet.Rows.Count > i; i++)
            {
                if (CheckRowDataVailabile(ThreadSahdeDataSheet.Rows[i]))
                {
                    ThreadShadeModel model = new ThreadShadeModel()
                    {
                        Index = i,
                        Season = ThreadSahdeDataSheet.Rows[i].Cells[0].ColValue,
                        FS = ThreadSahdeDataSheet.Rows[i].Cells[1].ColValue,
                        WashOrNonWash = ThreadSahdeDataSheet.Rows[i].Cells[2].ColValue,
                        FabricReference = ThreadSahdeDataSheet.Rows[i].Cells[3].ColValue,
                        RMBaseColorNameAndCode = ThreadSahdeDataSheet.Rows[i].Cells[4].ColValue,
                        WashTechAndColor = ThreadSahdeDataSheet.Rows[i].Cells[5].ColValue,
                        RequestedThread = ThreadSahdeDataSheet.Rows[i].Cells[6].ColValue,
                        ThreadSahde = ThreadSahdeDataSheet.Rows[i].Cells[7].ColValue,
                        RepairThreadShade = ThreadSahdeDataSheet.Rows[i].Cells[8].ColValue,
                        Comment = ThreadSahdeDataSheet.Rows[i].Cells[9].ColValue,

                    };
                    result.Add(model);
                }
            }
            return result;
        }

        private List<LabdipChartModel> SetLabdipThreadShadeProcessData(List<LabdipChartModel> labdipList, List<ThreadShadeModel> threadShadeList, List<string> threadTypes)
        {
            List<LabdipChartModel> newLabdipRecordCollection = new List<LabdipChartModel>();

            threadShadeList.ForEach(threadEle =>
            {
                var foundList = labdipList.FindAll(ele => ele.ItemName.Contains(threadEle.FabricReference) && ele.RMColor.Equals(threadEle.RMBaseColorNameAndCode));
                if (foundList != null)
                {
                    foundList.ForEach(ele =>
                    {
                        PrepareNewRecord(ele, threadEle, threadTypes).ForEach(newEle =>
                        {
                            newLabdipRecordCollection.Add(newEle);
                        });
                    });
                }

            });

            labdipList.AddRange(newLabdipRecordCollection);
            labdipList.Sort((idx1, idx2) => idx1.Index.CompareTo(idx2.Index));
            //before return sort the list by Index Number
            return labdipList;
        }

        private bool CheckRowDataVailabile(Row row)
        {
            return (row.Cells.Find(c => c.ColValue != null && c.ColValue.Length > 0) == null) ? false : true;
        }

        private int StringToInt(string anyString)
        {
            int number = -1;
            bool isParsable = Int32.TryParse(anyString, out number);
            return number;
        }

        private List<LabdipChartModel> PrepareNewRecord(LabdipChartModel labdip, ThreadShadeModel thread, List<string> threadTypes)
        {
            List<LabdipChartModel> result = new List<LabdipChartModel>();

            threadTypes.ForEach(ele =>
            {
                LabdipChartModel labdipChartModel = new LabdipChartModel()
                {
                    Index = labdip.Index,
                    Division = labdip.Division,
                    Season = labdip.Season,
                    Category = labdip.Category,
                    Program = labdip.Program,
                    StyleNoIndividual = labdip.StyleNoIndividual,
                    GMTColor = labdip.GMTColor,
                    GMTDescription = labdip.GMTDescription,
                    NRF = labdip.NRF,
                    ColorCode = "",
                    RMColor = thread.ThreadSahde,
                    PackCombination = labdip.PackCombination,
                    ColorDyeingTechnic = labdip.ColorDyeingTechnic,
                    PalcementName = "",
                    BOMSelection = labdip.BOMSelection,
                    ItemName = GetThredTypeAndTicket(ele),
                    SupplierName = GetThredTypeSupplier(ele),
                    RMColorRef = thread.RMBaseColorNameAndCode,
                    GarmentWay = labdip.GarmentWay,
                    FBNumber = labdip.FBNumber,
                    MaterialType = labdip.MaterialType,
                };
                result.Add(labdipChartModel);
            });
            return result;
        }

        private string GetThredTypeAndTicket(string thredType)
        {
            string result = string.Empty;
            _thredTypeList.ForEach(ele =>
            {
                if (ele.ThreadType.Equals(thredType))
                {
                    result = ele.ThreadType + " " + ele.TicketNo;
                }
            });
            return result;
        }

        private string GetThredTypeSupplier(string thredType)
        {
            return _thredTypeList.Find(ele => ele.ThreadType.Equals(thredType)).Supplier;
        }
        #endregion


    }
}
