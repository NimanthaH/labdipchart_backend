using BrandixAutomation.Labdip.API.Models;
using ExcelDataReader;
using log4net;
using log4net.Config;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace BrandixAutomation.Labdip.API.ProcessFiles
{
    public class LabdipChartDataService
    {
        #region Variable Declarations

        ILog logger = LogManager.GetLogger(typeof(LabdipChartDataService));

        private LabdipChartModel _labdipChartModel;
        private List<LabdipChartModel> _labdipChartModelList;
        private List<LabdipChartModel> _labdipChartModelData;
        private Sheet _sheet;
        private Sheet _linematrixsheet;
        private LabdipChartModel _labdipChartHeaders;
        private string _error;

        #endregion

        #region Constructor
        public LabdipChartDataService()
        {
            _labdipChartModel = new LabdipChartModel();
            _labdipChartModelList = new List<LabdipChartModel>();
            _labdipChartModelData = new List<LabdipChartModel>();
            _labdipChartHeaders = new LabdipChartModel();
            _sheet = new Sheet("MasterTemp");
            _linematrixsheet = new Sheet("MasterLineMatrix");

            //Create log4net reference
            var logRepository = LogManager.GetRepository(Assembly.GetEntryAssembly());
            XmlConfigurator.Configure(logRepository, new FileInfo("log4net.config"));
            logger = LogManager.GetLogger(typeof(Program));
        }

        #endregion

        #region APIs
        public List<LabdipChartModel> GetLabdipChartData(IFormFileCollection response_files, ICollection<string> keys) //byte[] excelSheetByteArray
        {
            logger.InfoFormat("GetLabdipChartData function called with response_files={0}, keys={1}", response_files, keys);

            try
            {
                //get files
                var line_matrix = response_files.Where(c => c.Name.ToString() == "linematrixfile").FirstOrDefault();
                var tech_packs = response_files.Where(c => c.Name.ToString() == "techpackfile").ToList();

                string option = keys.ToArray()[0].ToString();
                int file_variation = Convert.ToInt16(keys.ToArray()[1].ToString());

                OptionService optionservice = new OptionService();
                String type = (option != null && option != "") ? optionservice.GetOptionById(option) != null ? optionservice.GetOptionById(option).Type : "NF" : "NF";

                var b = new List<LabdipChartModel>();

                //loop through tech pack files
                if (tech_packs != null && tech_packs.Count > 0)
                {

                    foreach (var tech_pack in tech_packs)
                    {
                        if (tech_pack != null && tech_pack.Length > 0)
                        {
                            //Call the labdip process
                            if (type == "TP")
                            {
                                _labdipChartModelData.AddRange(ProcessLabdipChartDataPDF(tech_pack));
                            }
                            else if (type == "BOM" && line_matrix != null && line_matrix.Length > 0)
                            {
                                _labdipChartModelData.AddRange(ProcessLabdipChartDataExcel(tech_pack, line_matrix, file_variation));
                            }
                        }

                    }
                }
            }
            catch (Exception error)
            {
                _error = _error + error.Message + ". ";
                return null;
            }

            return _labdipChartModelData;
        }

        #region process_pdf

        public List<LabdipChartModel> ProcessLabdipChartDataPDF(IFormFile file) //byte[] excelSheetByteArray
        {
            _sheet = ReadDataIntoSheet(file, false, "");

            //1 Set Divion
            SetDivision();

            //2 Season and Category
            Season_Category();

            //3 Styele No and DEscription
            StyleNoIndividual_GMTDescription();

            //4 GMTColor RMColor
            _labdipChartModelList = PrepareLabdipChart(ColorwayPivotSheet(GMT_Color_RMColor_ModelSearch(_sheet), GetSectionDetails(_sheet)));

            PrepareNRFandColorCode();

            SetColorDyeingTechnic();

            //Update error
            _labdipChartModelList.ForEach(c => c.error = _error);

            return _labdipChartModelList;
        }

        private Sheet ReadDataIntoSheet(IFormFile file, Boolean filterData, string filterValue) //byte[] excelSheetByteArray
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
                        Boolean valid = !filterData;
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            Cell cell = PrepareCell(r, i, reader.GetValue(i));
                            if (filterData && filterValue.IndexOf(cell.ColValue) < 0)
                            {
                                valid = false;
                            }

                            row.Cells.Add(cell);
                        }
                        r++;
                        if (valid) { sheet.Rows.Add(row); }
                    }
                    reader.Close(); //only one Sheet reading..
                } while (reader.NextResult());
            }
            //}
            return sheet;
        }

        private Cell PrepareCell(int rowindex, int colindex, object cellValue)
        {
            return new Cell()
            {
                ColIndex = colindex,
                RowIndex = rowindex,
                ColValue = cellValue != null ? cellValue.ToString() : "Null"
            };
        }

        private void SetDivision()
        {
            bool found = false;

            _sheet.Rows.ForEach(row =>
            {
                row.Cells.ForEach(cell =>
                {
                    if ((found) && (cell.ColValue != null) && (!cell.ColValue.Equals("Null")))
                    {
                        _labdipChartModel.Division = cell.ColValue;
                        found = false;
                    }

                    if ((cell.ColValue != null) && (cell.ColValue.Equals("Sub - Brand")))
                        found = true;
                });
            });
        }

        private void Season_Category()
        {
            _sheet.Rows.ForEach(row =>
            {
                row.Cells.ForEach((cell) =>
                {
                    if ((cell.ColValue != null) && cell.ColValue.Contains("Season Name"))
                    {
                        //int i = 0;
                        string result_season = "";
                        string result_category = "";

                        var temp = cell.ColValue.Split(' ');

                        bool found = false;
                        if (temp != null && temp.Length > 0)
                        {
                            for (int i = 0; i < temp.Length; i++)
                            {
                                if (found == false && temp[(i + 1)] != null && temp[i].Contains("Season") & temp[(i + 1)].Contains("Name"))
                                {
                                    result_season = temp[i + 2] + temp[i + 3];
                                    result_category = temp[i + 4] + temp[i + 5];
                                    found = true;
                                }
                            }
                        }

                        _labdipChartModel.Season = result_season;
                        _labdipChartModel.Category = result_category;
                    }
                });
            });
        }

        private void StyleNoIndividual_GMTDescription()
        {
            bool found = false;

            _sheet.Rows.ForEach((row) =>
            {
                row.Cells.ForEach(cell =>
                {
                    if ((found) && (cell.ColValue != null) && (!cell.ColValue.Equals("Null")))
                    {
                        _labdipChartModel.StyleNoIndividual = cell.ColValue.Substring(0, cell.ColValue.IndexOf('/'));
                        _labdipChartModel.GMTDescription = cell.ColValue.Substring(cell.ColValue.IndexOf("/") + 1); //each.Substring(each.IndexOf("/")+1, each.IndexOf('/')-2);
                        found = false;
                    }

                    if ((cell.ColValue != null) && (cell.ColValue.Equals("Product Code / Description:")))
                    {
                        found = true;
                    }
                });
            });
        }

        public List<Sheet> GMT_Color_RMColor_ModelSearch(Sheet tempSheet)
        {
            bool found = false;
            int colorwayHeaderFoundColIndex = 0;
            int colorwayHeaderFoundRowIndex = 0;
            int colorwayHeaderFoundColFrequency = 0;
            Sheet colorWaySheet = new Sheet("Temp");
            List<Sheet> colorWayExtractedList = new List<Sheet>();

            for (int r = 0; r < tempSheet.Rows.Count; r++)
            {
                Row colorwayRow = new Row();
                foreach (Cell cell in tempSheet.Rows[r].Cells)
                {
                    if (found)
                    {
                        if ((cell.ColIndex >= colorwayHeaderFoundColIndex) && (cell.ColIndex <= (colorwayHeaderFoundColIndex + 30)) && (r > colorwayHeaderFoundRowIndex))
                        {
                            if ((colorwayRow.Cells.Count < 10) && (ValidateColumnData(tempSheet.Rows[r], cell, colorWaySheet.Rows.Count > 0 ? colorWaySheet.Rows[0] : null, tempSheet.Rows[colorwayHeaderFoundRowIndex + 1])))
                            {
                                if ((cell.ColValue != null) && (!cell.ColValue.Equals("Null")))
                                {
                                    var cellHeader = new Cell() { ColIndex = cell.ColIndex, ColValue = cell.ColValue.Replace("\n", "").Replace("\r", "") };
                                    colorwayRow.Cells.Add(cellHeader);
                                }
                            }
                        }
                    }

                    if ((cell.ColValue != null) && cell.ColValue.ToLower().Equals("colorway") || (r == (tempSheet.Rows.Count - 1)))
                    {
                        found = true;
                        if (colorWaySheet.Rows.Count > 0)
                        {
                            Sheet newSheet = new Sheet("Colorway");
                            newSheet.Rows.AddRange(colorWaySheet.Rows);
                            colorWayExtractedList.Add(newSheet);
                            colorWaySheet.Rows.Clear();
                        }
                        else
                        {
                            if (colorwayHeaderFoundColFrequency > 0)
                            {
                                Sheet newSheet = new Sheet("Colorway");
                                colorWayExtractedList.Add(newSheet);
                                colorWaySheet.Rows.Clear();
                            }

                        }
                        colorwayHeaderFoundColIndex = cell.ColIndex;
                        colorwayHeaderFoundRowIndex = r;
                        colorwayHeaderFoundColFrequency++;
                        break;
                    }
                };
                if (colorwayRow.Cells.Count > 0)
                    colorWaySheet.Rows.Add(colorwayRow);
            };
            return colorWayExtractedList;
        }

        private List<Sheet> GetSectionDetails(Sheet tempSheet)
        {
            bool found = false;
            int sectioneaderFoundColIndex = 0;
            int sectionHeaderFoundRowIndex = 0;
            Sheet sectionSheet = new Sheet("Temp");
            List<Sheet> sectionExtractedList = new List<Sheet>();
            for (int r = 0; r < tempSheet.Rows.Count; r++)
            {
                Row sectionRow = new Row();
                foreach (Cell cell in tempSheet.Rows[r].Cells)
                {
                    if (found)
                    {
                        if ((cell.ColIndex >= sectioneaderFoundColIndex) && (cell.ColIndex <= (sectioneaderFoundColIndex + 30)) && (r > sectionHeaderFoundRowIndex))
                        {
                            if ((sectionRow.Cells.Count < 7) && (ValidateColumnData(tempSheet.Rows[r], cell, sectionSheet.Rows.Count > 0 ? sectionSheet.Rows[0] : null, tempSheet.Rows[sectionHeaderFoundRowIndex + 1])))
                            {
                                if (sectionSheet.Rows.Count > 0) //Column Data
                                {
                                    if ((cell.ColValue != null))
                                    {
                                        var cellHeader = new Cell() { ColIndex = cell.ColIndex, ColValue = cell.ColValue.Replace("\n", "").Replace("\r", "") };
                                        sectionRow.Cells.Add(cellHeader);
                                    }
                                }
                                else //Headers reading
                                {
                                    if ((cell.ColValue != null) && (!cell.ColValue.Equals("Null")))
                                    {
                                        var cellHeader = new Cell() { ColIndex = cell.ColIndex, ColValue = cell.ColValue.Replace("\n", "").Replace("\r", "") };
                                        sectionRow.Cells.Add(cellHeader);
                                    }
                                }
                            }
                        }
                    }


                    if ((cell.ColValue != null) && cell.ColValue.ToLower().Contains("section:") || (r == (tempSheet.Rows.Count - 1)))
                    {
                        found = true;
                        //First sheet detection
                        if (sectionSheet.Rows.Count == 0) sectionSheet.SheetName = cell.ColValue.Substring(cell.ColValue.IndexOf(":") + 1);

                        if (sectionSheet.Rows.Count > 0)
                        {
                            Sheet newSheet = new Sheet(sectionSheet.SheetName);
                            newSheet.Rows.AddRange(sectionSheet.Rows);
                            sectionExtractedList.Add(newSheet);
                            sectionSheet.Rows.Clear();
                            sectionSheet.SheetName = cell.ColValue.Substring(cell.ColValue.IndexOf(":") + 1);
                        }
                        sectioneaderFoundColIndex = cell.ColIndex;
                        sectionHeaderFoundRowIndex = r;
                        break;
                    }
                }
                if (sectionRow.Cells.Count > 0)
                    sectionSheet.Rows.Add(sectionRow);
            }
            return sectionExtractedList;
        }

        private void PrepareNRFandColorCode()
        {
            _labdipChartModelList.ForEach(each =>
            {
                if (each.GMTColor != null)
                {
                    var indexLocation = each.GMTColor.LastIndexOf("(");
                    var endTale = each.GMTColor.Substring(indexLocation - 5);
                    var nrf = endTale.Substring(0, 5);
                    each.NRF = nrf;
                }
                else
                    each.NRF = null;


                //setting color codes
                // same thing can get from ection data area (part Type Col)
                if (each.RMColor != null)
                {
                    string colorCode = null;
                    if (each.BOMSelection.ToLower().Equals("fabric"))
                    {
                        var wordArray = each.RMColor.Split(' ');
                        if (wordArray.Length > 1)
                        {
                            foreach (var word in wordArray)
                            {
                                colorCode = (word.Length == 4) ? Regex.IsMatch(word, @"^[a-zA-Z0-9]+$") ? word : null : null; //Regex reemove non leter and number fields, 
                            }
                        }
                    }
                    each.ColorCode = colorCode;
                }
                else
                    each.ColorCode = null;

            });
        }

        private void SetColorDyeingTechnic()
        {
            _labdipChartModelList.ForEach((each) =>
            {
                if ((each.PalcementName != null) && (each.BOMSelection != null) && (each.ItemName != null))
                {
                    if (each.BOMSelection.Contains("Fabric"))
                    {
                        if (each.Index < _labdipChartModelList.Count)
                        {
                            if ((_labdipChartModelList[each.Index].BOMSelection != null) && (_labdipChartModelList[each.Index].PalcementName != null))
                            {
                                if (_labdipChartModelList[each.Index].BOMSelection.ToLower().Contains("fabric") && _labdipChartModelList[each.Index].PalcementName.ToLower().Contains("dye"))
                                {
                                    each.ColorDyeingTechnic = _labdipChartModelList[each.Index].RMColor;
                                }
                            }
                        }
                    }
                }
            });
        }

        private bool ValidateColumnData(Row row, Cell currentCell, Row firstRow = null, Row sectionHeaderRow = null)
        {
            bool result = false;
            row.Cells.ForEach(cell =>
            {
                //if ((cell.ColValue != null) && ((cell.ColValue.Equals("Quantity")) || cell.ColValue.Equals("0")) && (cell.ColIndex == 52))
                if ((cell.ColValue != null) && (cell.ColValue.ToLower().Contains("quantity") || (cell.ColValue.ToLower().Contains("price")) || (decimal.TryParse(cell.ColValue, out decimal cellval))))
                {
                    if (firstRow == null)
                    {
                        result = ValidateColorwayHeader(currentCell.ColValue) || ValidateSectionHeaderData(currentCell.ColValue) ? true : false;
                    }
                    else
                    {
                        result = currentCell.ColIndex < cell.ColIndex ? ColumAvailability(firstRow, currentCell, row, sectionHeaderRow) : false;
                    }
                }
            });
            return result;
        }

        private bool ValidateColorwayHeader(string cellContent)
        {
            string patternText = @"[(][0-9]{8}[)]$";
            Regex reg = new Regex(patternText);
            return reg.IsMatch(cellContent) ? true : false;
        }

        private bool ValidateSectionHeaderData(string cellContent)
        {
            List<string> columnHeaderNames = new List<string>(new string[] { "Part Type", "Part Name", "Material Id", "Material", "Over-ride", "Supplier Quality Number", "Supplier", "Use (Placement)" });

            return columnHeaderNames.Contains(cellContent) ? true : false;
        }

        private bool ColumAvailability(Row firstRow, Cell currentCell, Row currentRow, Row sectionHeaderRow)
        {
            if (firstRow == null) return true;
            bool result = false;

            if (firstRow.Cells.FindIndex(ele => ele.ColIndex == currentCell.ColIndex) != -1)
            {
                var firstRowEndIndex = sectionHeaderRow.Cells.FindIndex(ele => ele.ColValue.ToLower().Contains("quantity") || ele.ColValue.ToLower().Contains("price"));
                if (firstRowEndIndex > 0)
                {
                    if ((decimal.TryParse(currentRow.Cells[firstRowEndIndex].ColValue, out decimal cellval)))
                        result = true;
                }
            }

            return result;
        }

        private Sheet ColorwayPivotSheet(List<Sheet> colorwayExtractedData, List<Sheet> sectionExtractedData)
        {
            Sheet colorwayPivotSheet = new Sheet("ColorwayPivot");
            int sectionSheetNo = 0;

            colorwayExtractedData.ForEach(sheet =>
            {
                if (sheet.Rows.Count > 0)
                {
                    //column count
                    for (int c = 0; c < sheet.Rows[0].Cells.Count; c++)
                    {
                        //row by row
                        for (int r = 1; r < sheet.Rows.Count; r++)
                        {
                            Cell GMTColor = new Cell() { ColIndex = 0, ColValue = sheet.Rows[0].Cells[c].ColValue };
                            Cell RMColor = new Cell() { ColIndex = r, ColValue = sheet.Rows[r].Cells[c].ColValue };
                            Cell partName = new Cell() { ColIndex = 2, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Part Name", "Part Type" })) };
                            Cell materialId = new Cell() { ColIndex = 3, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Material Id" })) };
                            Cell material = new Cell() { ColIndex = 4, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Material" })) };
                            Cell supplierQNo = new Cell() { ColIndex = 5, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Supplier Quality Number" })) };
                            Cell supplier = new Cell() { ColIndex = 6, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Supplier" })) };
                            Cell usePalcemant = new Cell() { ColIndex = 7, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Use (Placement)" })) };
                            Cell section = new Cell() { ColIndex = 8, ColValue = NullReplacement(sectionExtractedData[sectionSheetNo].SheetName) };

                            Row row = new Row();
                            row.Cells.Add(GMTColor);
                            row.Cells.Add(RMColor);
                            row.Cells.Add(partName);
                            row.Cells.Add(materialId);
                            row.Cells.Add(material);
                            row.Cells.Add(supplierQNo);
                            row.Cells.Add(supplier);
                            row.Cells.Add(usePalcemant);
                            row.Cells.Add(section);
                            colorwayPivotSheet.Rows.Add(row);
                        }
                    }
                }
                else
                {
                    if (sectionExtractedData[sectionSheetNo].Rows.Count > 1)
                    {
                        for (int r = 1; r < sectionExtractedData[sectionSheetNo].Rows.Count; r++)
                        {
                            Cell GMTColor = new Cell() { ColIndex = 0, ColValue = null };
                            Cell RMColor = new Cell() { ColIndex = 1, ColValue = null };
                            Cell partName = new Cell() { ColIndex = 2, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Part Name", "Part Type" })) };
                            Cell materialId = new Cell() { ColIndex = 3, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Material Id" })) };
                            Cell material = new Cell() { ColIndex = 4, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Material" })) };
                            Cell supplierQNo = new Cell() { ColIndex = 5, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Supplier Quality Number" })) };
                            Cell supplier = new Cell() { ColIndex = 6, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Supplier" })) };
                            Cell usePalcemant = new Cell() { ColIndex = 7, ColValue = NullReplacement(SetSectionData(sectionExtractedData[sectionSheetNo].Rows[0], sectionExtractedData[sectionSheetNo].Rows[r], new string[] { "Use (Placement)" })) };
                            Cell section = new Cell() { ColIndex = 8, ColValue = NullReplacement(sectionExtractedData[sectionSheetNo].SheetName) };

                            Row row = new Row();
                            row.Cells.Add(GMTColor);
                            row.Cells.Add(RMColor);
                            row.Cells.Add(partName);
                            row.Cells.Add(materialId);
                            row.Cells.Add(material);
                            row.Cells.Add(supplierQNo);
                            row.Cells.Add(supplier);
                            row.Cells.Add(usePalcemant);
                            row.Cells.Add(section);
                            colorwayPivotSheet.Rows.Add(row);
                        }
                    }
                }
                sectionSheetNo++;
            });

            return colorwayPivotSheet;
        }

        private List<LabdipChartModel> PrepareLabdipChart(Sheet colorwayPivotSheet)
        {
            List<LabdipChartModel> labdipChartModels = new List<LabdipChartModel>();
            int index = 1;
            colorwayPivotSheet.Rows.ForEach(row =>
            {
                LabdipChartModel newlabDipChartModel = new LabdipChartModel()
                {
                    Index = index,
                    Division = _labdipChartModel.Division,
                    Season = _labdipChartModel.Season,
                    Category = _labdipChartModel.Category,
                    Program = null,
                    PackCombination = null,
                    StyleNoIndividual = _labdipChartModel.StyleNoIndividual,
                    GMTDescription = _labdipChartModel.GMTDescription,
                    GMTColor = row.Cells[0].ColValue,
                    RMColor = row.Cells[1].ColValue,
                    PalcementName = row.Cells[7].ColValue,
                    BOMSelection = row.Cells[8].ColValue,
                    ItemName = row.Cells[5].ColValue,
                    SupplierName = row.Cells[6].ColValue,
                    MaterialType = row.Cells[4].ColValue,
                    FBNumber = ExtractFBNumber(row.Cells[4].ColValue), //row.Cells[3].ColValue 
                    GarmentWay = FilterGarmentWay(row.Cells[2].ColValue, row.Cells[4].ColValue, row.Cells[1].ColValue, labdipChartModels, index, row.Cells[7].ColValue)
                };
                labdipChartModels.Add(newlabDipChartModel);
                index++;
            });
            return labdipChartModels;
        }

        private byte[] FileToByteArray(string fileName)
        {
            byte[] fileContent = null;
            System.IO.FileStream fs = new System.IO.FileStream(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            System.IO.BinaryReader binaryReader = new System.IO.BinaryReader(fs);
            long byteLength = new System.IO.FileInfo(fileName).Length;
            fileContent = binaryReader.ReadBytes((Int32)byteLength);
            fs.Close();
            fs.Dispose();
            binaryReader.Close();
            return fileContent;
        }

        private string NullReplacement(string stringVal)
        {
            return string.IsNullOrEmpty(stringVal) ? null : stringVal.Equals("Null") ? null : stringVal;
        }

        private string SetSectionData(Row headerRow, Row currentRow, string[] headerNameList)
        {
            string result = null;

            for (int i = 0; i < headerRow.Cells.Count; i++)
            {
                if (Array.Exists(headerNameList, s => s.Equals(headerRow.Cells[i].ColValue)))
                    result = currentRow.Cells[i].ColValue;
            }
            return result;
        }

        private string ExtractFBNumber(string material)
        {
            var stack = new Stack<char>();

            foreach (var c in ReverseString(material))
            {
                if (!char.IsDigit(c))
                    break;
                stack.Push(c);
            }

            return new string(stack.ToArray());
        }

        private string ReverseString(string stringInput)
        {
            // With Inbuilt Method Array.Reverse Method  
            char[] charArray = stringInput.ToCharArray();
            Array.Reverse(charArray);
            return (new string(charArray));
        }

        private string FilterGarmentWay(string partType, string material, string colorWayCol, List<LabdipChartModel> labdipChartModels, int currentIndex, string placementName)
        {
            string result = "";
            if (!string.IsNullOrEmpty(material) && !string.IsNullOrEmpty(partType))
                result = (partType.ToLower().Contains("comments") && material.ToLower().Contains("print")) ? UpdatePreviousGarmentwayColumns(currentIndex,
                                                                                                                                             labdipChartModels,
                                                                                                                                             placementName,
                                                                                                                                             colorWayCol) : "";
            return result;
        }

        private string UpdatePreviousGarmentwayColumns(int index, List<LabdipChartModel> labdipChartModels, string placementName, string colorWay)
        {
            bool notFound = true;
            int rowIndex = index - 1;
            while (notFound && rowIndex > 0)
            {
                rowIndex--;
                if (labdipChartModels[rowIndex].PalcementName.Equals(placementName))
                {
                    labdipChartModels[rowIndex].GarmentWay = colorWay;
                }
                else
                {
                    notFound = false;
                }
            }
            return colorWay;
        }

        #endregion

        #region process_excel
        //Option Excel BOM With Line Matrix
        public List<LabdipChartModel> ProcessLabdipChartDataExcel(IFormFile tech_pack, IFormFile linematrix, int file_variation) //byte[] excelSheetByteArray
        {
            logger.InfoFormat("ProcessLabdipChartDataExcel function called with tech_pack={0}, linematrix={1}, file_variation={2}", tech_pack, linematrix, file_variation);
            //Init variables
            List<LabdipChartModel> labdipChartModelList = new List<LabdipChartModel>();

            //1 Read and set sheet data
            _linematrixsheet = ReadDataIntoSheet(linematrix, false, "");
            _sheet = ReadDataIntoSheet(tech_pack, false, "");

            //2 Set Header Details - BOM
            labdipChartModelList = processFileData(tech_pack.FileName, "bom", file_variation, _sheet);

            //3 Arrange sub columns data
            labdipChartModelList = arrangeSubColumns(labdipChartModelList);
            _labdipChartModelList = labdipChartModelList;

            //4 Set Header Details - line matrix
            processFileData(tech_pack.FileName, "line_matrix", file_variation, _linematrixsheet);

            //5 SpliT and Transform Data
            _labdipChartModelList = splitTransformData(_labdipChartModelList, file_variation);

            return _labdipChartModelList.Where(c => c.ItemName != null && c.ItemName != "" && c.ItemName != "Null").ToList();
        }

        #region: step - process data to headers
        //Process the sheet data according to the header master and create a lab dip chart list
        private List<LabdipChartModel> processFileData(string fileName, string file_type, int file_variation, Sheet sheet)
        {
            logger.InfoFormat("processFileData function called with filename={0}, type={1}, file_variation={2}", fileName, file_type, file_variation);

            //Initlaize variables
            HeaderMasterService headerMasterService = new HeaderMasterService();

            List<LabdipChartModel> labdipChartModelList = new List<LabdipChartModel>();
            List<HeaderMaster> _headers = headerMasterService.GetHeaderList();
            List<HeaderMaster> _typewise_headers = new List<HeaderMaster>();
            List<HeaderMaster> headers = new List<HeaderMaster>();
            LabdipChartModel repeatData = new LabdipChartModel();

            try
            {
                //Set Hedaer Columns
                if (_headers != null && _headers.Count > 0)
                {
                    //Get headers for file type and customer variation
                    _typewise_headers = _headers.Where(c => c.FileName == file_type && c.Variation == file_variation).ToList();

                    if (_typewise_headers != null && _typewise_headers.Count > 0)
                    {
                        //Repeat through rows
                        sheet.Rows.ForEach(row =>
                        {   
                            //Initlaize variable for row data
                            LabdipChartModel labdipChartRow = new LabdipChartModel();
                            Boolean validRow = true;
                            //Initlaize sub columns
                            labdipChartRow.SubColumns = new List<LabdipChartSubModel>();

                            //Repeat through cells
                            row.Cells.ForEach(cell =>
                            {
                                //Check for nullability of cell
                                if (cell != null)
                                {

                                    //Set Row index
                                    labdipChartRow.RowIndex = cell.RowIndex;

                                    #region: Static header process

                                    //Header catch flag
                                    Boolean isStaticHeader = false;

                                    //Start: Header setting
                                    //Check for the validity of the cell value
                                    if (cell.ColValue != null && cell.ColValue.Trim() != "")
                                    {
                                        //Check wether is this a header cell according to cell value and header master header name and customer variation
                                        List<HeaderMaster> staticColumnHeaders = _typewise_headers.Where(c => c.HedaerName.ToString().Trim() == cell.ColValue.ToString().Trim() && c.HeaderType == "static").ToList();
                                        //Check is this a header cell
                                        if (staticColumnHeaders != null && staticColumnHeaders.Count > 0)
                                        {
                                            //loop through all identified headers to update reference
                                            foreach (HeaderMaster staticColumnHeader in staticColumnHeaders)
                                            {
                                                if (staticColumnHeader != null)
                                                {
                                                    //Set row and column of header cell
                                                    staticColumnHeader.RowNo = cell.RowIndex;
                                                    staticColumnHeader.ColumnNo = cell.ColIndex;
                                                    //Update header arrays
                                                    headers.Add(staticColumnHeader);
                                                    isStaticHeader = true;
                                                }
                                            }
                                        }
                                    }
                                    //End: Header setting

                                    //Start: Data Setting
                                    //Check is there any headers identified as at now 
                                    if (headers != null && headers.Count > 0 && !isStaticHeader)
                                    {
                                        //Check for cell validity of cell to and header attributes
                                        List<HeaderMaster> cellHeaderRefernces = headers.Where(c => c.ColumnNo + c.ColumnSpan == cell.ColIndex && c.HeaderType == "static").ToList();
                                        //Loop through header references
                                        foreach (HeaderMaster cellHeaderRefernce in cellHeaderRefernces)
                                        {
                                            //Check for Row reference match with span
                                            if (cellHeaderRefernce != null)
                                            {
                                                //Check for filter data
                                                if (cellHeaderRefernce.FilterData != null && cellHeaderRefernce.FilterData != "no")
                                                {
                                                    //Check is a the cell value contains in the file name
                                                    int index = cell.ColValue != null ? (cellHeaderRefernce.FilterData == "filename" ? fileName.IndexOf(cell.ColValue.ToString().Trim()) : cellHeaderRefernce.FilterData.IndexOf(cell.ColValue.ToString().Trim())) : -1;

                                                    //Checker for filter
                                                    if (validRow && index < 0)
                                                    {
                                                        validRow = false; //Set flag to not add the row and update data
                                                    }
                                                }

                                                //Skip updating the rows which a value is mandatory and connot be blank (Garament to RM color mapping)
                                                if (cellHeaderRefernce.SkipBlanks && (cell.ColValue != null || cell.ColValue.ToString().Trim() == "" || cell.ColValue.ToString().Trim() == "Null"))
                                                {
                                                    validRow = false; //Set flag to not add the row and update data
                                                }

                                                //Get property of matching model attribute to fill data
                                                var propertyInfo = typeof(LabdipChartModel).GetProperty(cellHeaderRefernce.HeaderAttribute);
                                                //Start: Value update and setting normal flow
                                                if (cell.ColValue != null && propertyInfo != null && validRow)
                                                {
                                                    //Check is the data need to repeat for all cells
                                                    if (cellHeaderRefernce.UpdateforAll)
                                                    {
                                                        //Update all data with relevant cell value
                                                        _labdipChartModelList.ForEach(labdipChartModel => propertyInfo.SetValue(labdipChartModel, Convert.ChangeType(processCellValue(cell.ColValue, cellHeaderRefernce), propertyInfo.PropertyType)));
                                                    }
                                                    else
                                                    {
                                                        //Update Row data with relevant cell value
                                                        propertyInfo.SetValue(labdipChartRow, Convert.ChangeType(processCellValue(cell.ColValue, cellHeaderRefernce), propertyInfo.PropertyType));

                                                        //Check for repeat data updating
                                                        if (cellHeaderRefernce.RepeatData && repeatDataValueValidator(cell.ColValue))
                                                        {
                                                            //Set latest value as repeat data property
                                                            propertyInfo.SetValue(repeatData, Convert.ChangeType(processCellValue(cell.ColValue, cellHeaderRefernce), propertyInfo.PropertyType));
                                                        }
                                                    }

                                                }
                                                //End: Value update and setting normal flow

                                                //Start: Value update and setting repeat data
                                                if (propertyInfo != null && cellHeaderRefernce.RepeatData)
                                                {
                                                    //Get model property infomation to pick the data from repeat data list
                                                    var repeatPropertyInfo = repeatData.GetType().GetProperty(cellHeaderRefernce.HeaderAttribute, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);
                                                    if (repeatPropertyInfo != null)
                                                    {
                                                        //Update Row data with relevant repeat data value
                                                        propertyInfo.SetValue(labdipChartRow, Convert.ChangeType(repeatPropertyInfo.GetValue(repeatData, null), propertyInfo.PropertyType));
                                                    }

                                                }
                                                //End: Value update and setting repeat data
                                            }

                                        }
                                    }
                                    //End: Data Setting
                                }
                                #endregion

                                #region: Dynamic header process

                                //Header catch flag
                                Boolean isDynamicHeader = false;

                                //Start: Header setting
                                //Check for the validity of the cell value
                                if (cell.ColValue != null && cell.ColValue.Trim() != "")
                                {
                                    //Check wether is this a header cell according to cell value and header master header name and customer variation
                                    List<HeaderMaster> dynamicAllColumnHeaders = _typewise_headers.Where(c => c.HeaderType == "dynamic").ToList();
                                    //Check is this a header cell
                                    if (dynamicAllColumnHeaders != null && dynamicAllColumnHeaders.Count > 0)
                                    {
                                        List<HeaderMaster> dynamicBoundColumnHeaders = dynamicAllColumnHeaders.Where(c => c.LowerBoundHedaerName.ToString().Trim() == cell.ColValue.ToString().Trim() || c.UpperBoundHeaderName.ToString().Trim() == cell.ColValue.ToString().Trim()).ToList();
                                        //Check if the cell is an hedaer which lies around dynamic bounds
                                        if (dynamicBoundColumnHeaders != null && dynamicBoundColumnHeaders.Count > 0)
                                        {
                                            //loop through all identified dynamic headers to update reference
                                            foreach (HeaderMaster dynamicBoundColumnHeader in dynamicBoundColumnHeaders)
                                            {
                                                if (dynamicBoundColumnHeader != null)
                                                {
                                                    //Switch through update variations
                                                    switch (Tuple.Create(headers.Where(c => c.Id == dynamicBoundColumnHeader.Id).Count(), dynamicBoundColumnHeader.LowerBoundHedaerName, dynamicBoundColumnHeader.UpperBoundHeaderName))
                                                    {
                                                        //if header is already exsist, cell header is lower bound and upper bound is the end of the report
                                                        case var x when x.Item1 > 0 && x.Item2 == cell.ColValue.ToString().Trim() && x.Item3 != "END":
                                                            //Set row and column bounds of header cell
                                                            headers.Where(c => c.Id == dynamicBoundColumnHeader.Id).ToList().ForEach(c => c.LowerBoundRowNo = cell.RowIndex);
                                                            headers.Where(c => c.Id == dynamicBoundColumnHeader.Id).ToList().ForEach(c => c.LowerBoundColumnNo = cell.ColIndex);
                                                            break;
                                                        //if header is already exsist, cell header is lower bound and upper bound is not the end of the report
                                                        case var x when x.Item1 > 0 && x.Item2 == cell.ColValue.ToString().Trim() && x.Item3 == "END":
                                                            //Set row and column bounds of header cell
                                                            headers.Where(c => c.Id == dynamicBoundColumnHeader.Id).ToList().ForEach(c => c.LowerBoundRowNo = cell.RowIndex);
                                                            headers.Where(c => c.Id == dynamicBoundColumnHeader.Id).ToList().ForEach(c => c.LowerBoundColumnNo = cell.ColIndex);

                                                            //Set row and column bounds of header cell upper bound
                                                            headers.Where(c => c.Id == dynamicBoundColumnHeader.Id).ToList().ForEach(c => c.UpperBoundRowNo = cell.RowIndex);
                                                            headers.Where(c => c.Id == dynamicBoundColumnHeader.Id).ToList().ForEach(c => c.UpperBoundColumnNo = row.Cells.Count() + 1);
                                                            break;
                                                        //if header is not exsist, cell header is upper bound and upper bound is not the end of the report
                                                        case var x when x.Item1 > 0 && x.Item3 == cell.ColValue.ToString().Trim():
                                                            //Set row and column bounds of header cell
                                                            headers.Where(c => c.Id == dynamicBoundColumnHeader.Id).ToList().ForEach(c => c.UpperBoundRowNo = cell.RowIndex);
                                                            headers.Where(c => c.Id == dynamicBoundColumnHeader.Id).ToList().ForEach(c => c.UpperBoundColumnNo = cell.ColIndex);
                                                            break;
                                                        //if header is not exsist and cell header is lower bound
                                                        case var x when x.Item1 == 0 && x.Item2 == cell.ColValue.ToString().Trim() && x.Item3 != "END":
                                                            //Set row and column bounds of header cell
                                                            dynamicBoundColumnHeader.LowerBoundRowNo = cell.RowIndex;
                                                            dynamicBoundColumnHeader.LowerBoundColumnNo = cell.ColIndex;
                                                            //Update header arrays
                                                            headers.Add(dynamicBoundColumnHeader);
                                                            break;
                                                        //if header is not exsist and cell header is lower bound
                                                        case var x when x.Item1 == 0 && x.Item2 == cell.ColValue.ToString().Trim() && x.Item3 == "END":
                                                            //Set row and column bounds of header cell
                                                            dynamicBoundColumnHeader.LowerBoundRowNo = cell.RowIndex;
                                                            dynamicBoundColumnHeader.LowerBoundColumnNo = cell.ColIndex;

                                                            //Set row and column bounds of header cell upper bound
                                                            dynamicBoundColumnHeader.UpperBoundRowNo = cell.RowIndex;
                                                            dynamicBoundColumnHeader.UpperBoundColumnNo = row.Cells.Count() + 1;

                                                            //Update header arrays
                                                            headers.Add(dynamicBoundColumnHeader);
                                                            break;
                                                        //if header is not exsist and cell header is upper bound
                                                        case var x when x.Item1 == 0 && x.Item3 == cell.ColValue.ToString().Trim():
                                                            //Set row and column bounds of header cell
                                                            dynamicBoundColumnHeader.UpperBoundRowNo = cell.RowIndex;
                                                            dynamicBoundColumnHeader.UpperBoundColumnNo = cell.ColIndex;
                                                            //Update header arrays
                                                            headers.Add(dynamicBoundColumnHeader);
                                                            break;

                                                    }

                                                    isDynamicHeader = true;
                                                }
                                            }
                                        }
                                        //Not an dynamic bound column
                                        else
                                        {
                                            //Check wether the cell is on dynamic bound headers
                                            List<HeaderMaster> dynamicColumnHeaders = dynamicAllColumnHeaders.Where(c => c.LowerBoundRowNo == cell.RowIndex && c.UpperBoundRowNo <= cell.RowIndex && c.LowerBoundColumnNo < cell.ColIndex && c.UpperBoundColumnNo > cell.ColIndex).ToList();
                                            //found
                                            if (dynamicColumnHeaders != null && dynamicColumnHeaders.Count > 0)
                                            {
                                                //loop through all identified headers to update reference
                                                foreach (HeaderMaster dynamicColumnHeader in dynamicColumnHeaders)
                                                {
                                                    if (dynamicColumnHeader != null)
                                                    {
                                                        //Initilize new header (This is used because dynamicColumnHeader can;t update directly as it will update all simillar objects as well in the array)
                                                        HeaderMaster columnHeader = new HeaderMaster();
                                                        //Duplicate dynamicColumnHeader to columnHeader
                                                        columnHeader = headerMasterService.UpdateModelObject(dynamicColumnHeader);

                                                        //Set row and column of header cell
                                                        columnHeader.HedaerName = cell.ColValue;
                                                        columnHeader.RowNo = cell.RowIndex;
                                                        columnHeader.ColumnNo = cell.ColIndex;
                                                        //Update header arrays
                                                        headers.Add(columnHeader);

                                                        isDynamicHeader = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                //End: Header setting

                                //Start: Data Setting
                                //Check is there any headers identified as at now 
                                if (headers != null && headers.Count > 0 && !isDynamicHeader)
                                {
                                    //Check for cell validity of cell to and header attributes
                                    List<HeaderMaster> cellHeaderRefernces = headers.Where(c => c.ColumnNo + c.ColumnSpan == cell.ColIndex && c.HeaderType == "dynamic").ToList();
                                    //Loop through header references
                                    foreach (HeaderMaster cellHeaderRefernce in cellHeaderRefernces)
                                    {
                                        //Validate sub row addition
                                        Boolean validSubRow = true;
                                        //Check for Row reference match with span
                                        if (cellHeaderRefernce != null)
                                        {
                                            //Check for filter data
                                            if (cellHeaderRefernce.FilterData != null && cellHeaderRefernce.FilterData != "no")
                                            {
                                                //Check is a the cell value contains in the file name
                                                int index = cell.ColValue != null ? (cellHeaderRefernce.FilterData == "filename" ? fileName.IndexOf(cell.ColValue.ToString().Trim()) : cellHeaderRefernce.FilterData.IndexOf(cell.ColValue.ToString().Trim())) : -1;

                                                //Checker for filter
                                                if (validRow && index < 0)
                                                {
                                                    validSubRow = false; //Set flag to not add the row and update data
                                                }
                                            }

                                            //Skip updating the rows which a value is mandatory and connot be blank (Garament to RM color mapping)
                                            if (cellHeaderRefernce.SkipBlanks && (cell.ColValue != null || cell.ColValue.ToString().Trim() == "" || cell.ColValue.ToString().Trim() == "Null"))
                                            {
                                                validSubRow = false; //Set flag to not add the row and update data
                                            }

                                            //Get property of matching model attribute to fill data
                                            var propertyInfo = typeof(LabdipChartModel).GetProperty(cellHeaderRefernce.HeaderAttribute);
                                            //Start: Value update and setting normal flow
                                            if (cell.ColValue != null && propertyInfo != null)
                                            {
                                                //Update Row data with relevant cell value
                                                LabdipChartSubModel subColumns = new LabdipChartSubModel();

                                                //Update SubColumn data
                                                subColumns.Index = cell.RowIndex;
                                                subColumns.ColumnHeader = cellHeaderRefernce.HedaerName;
                                                subColumns.ColumnAttribute = cellHeaderRefernce.SubHeaderAttribute;
                                                subColumns.ValueAttribute = cellHeaderRefernce.SubValueAttribute;
                                                subColumns.ColumnValue = processCellValue(cellHeaderRefernce.SetSameHeaderasValue ? cellHeaderRefernce.HedaerName : cell.ColValue, cellHeaderRefernce);

                                                //Add sub column to row data
                                                if (validSubRow) { labdipChartRow.SubColumns.Add(subColumns); }

                                                //Check for repeat data updating
                                                if (cellHeaderRefernce.RepeatData && repeatDataValueValidator(cell.ColValue))
                                                {
                                                    //Set latest value as repeat data property
                                                    repeatData.SubColumns.Add(subColumns);
                                                }

                                            }
                                            //End: Value update and setting normal flow

                                            //Start: Value update and setting repeat data
                                            if (propertyInfo != null && cellHeaderRefernce.RepeatData)
                                            {
                                                //Get model property infomation to pick the data from repeat data list
                                                var repeatPropertyInfo = repeatData.GetType().GetProperty(cellHeaderRefernce.HeaderAttribute, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);
                                                if (repeatPropertyInfo != null)
                                                {
                                                    //Update Row data with relevant repeat data value
                                                    propertyInfo.SetValue(labdipChartRow, Convert.ChangeType(repeatPropertyInfo.GetValue(repeatData, null), propertyInfo.PropertyType));
                                                }

                                            }
                                            //End: Value update and setting repeat data
                                        }

                                    }
                                }
                                //End: Data Setting

                                #endregion
                            });

                            //Add row to lap dip chart list
                            if (validRow) { labdipChartModelList.Add(labdipChartRow); }

                        });
                    }
                }
                return labdipChartModelList;
            }
            catch (Exception error)
            {
                _error = _error + error.Message + ". ";
                logger.InfoFormat("error occured in processFileData function when called with filename={0}, type={1}, file_variation={2}, error={3}", fileName, file_type, file_variation, error.Message);
                return null;
            }
        }

        //Validate cell value to update repeat data array
        private Boolean repeatDataValueValidator(string cell_value)
        {
            Boolean proceed = true;

            switch (cell_value.Trim())
            {
                case null:
                    proceed = false;
                    break;
                case "":
                    proceed = false;
                    break;
            }

            return proceed;
        }

        //Process the cell value according to the header configs
        private string processCellValue(string cell_value, HeaderMaster config)
        {
            //Start: Split
            if (cell_value != null && config != null && config.Split != "no" && cell_value != "Null")
            {
                //debug
                if(cell_value.Contains("14-4201 TCX"))
                {
                    var t = true;
                }

                //Get all conditions as strings
                string[] settings = config.Split.Split("~");
                //Loop through splits
                foreach (string setting in settings)
                {
                    if (setting != null && setting != "")
                    {
                        //Split the conditions
                        string[] subSetting = setting.Split("|");
                        if (subSetting != null && subSetting.Count() > 0)
                        {
                            //Switch according to selection
                            switch (subSetting[0].ToString())
                            {
                                case "left":
                                    cell_value = cell_value.Length > Convert.ToInt16(subSetting[1]) ? cell_value.Substring(0, Convert.ToInt16(subSetting[1])) : cell_value;
                                    break;
                                case "right":
                                    cell_value = cell_value.Length > Convert.ToInt16(subSetting[1]) ? cell_value.Substring(cell_value.Length - Convert.ToInt16(subSetting[1]), Convert.ToInt16(subSetting[1])) : cell_value;
                                    break;
                                case "mid":
                                    cell_value = cell_value.Length > Convert.ToInt16(subSetting[1]) && cell_value.Length > Convert.ToInt16(subSetting[2]) && Convert.ToInt16(subSetting[2]) > Convert.ToInt16(subSetting[1]) ? cell_value.Substring(Convert.ToInt16(subSetting[1]), Convert.ToInt16(subSetting[2])) : cell_value;
                                    break;
                                case "delimited":
                                    if (subSetting[1] != null && subSetting[1] != "")
                                    {
                                        if (cell_value.IndexOf(subSetting[1]) >= 0)
                                        {
                                            if (subSetting[2] == "1")
                                            {
                                                cell_value = cell_value.Substring(0, cell_value.IndexOf(subSetting[1]));
                                            }
                                            else
                                            {
                                                cell_value = cell_value.Substring(cell_value.IndexOf(subSetting[1]), cell_value.Length - cell_value.IndexOf(subSetting[1]));
                                            }
                                        }
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
            //End: Split

            //Start: Replace
            if (config.Replace != "no")
            {
                //Get all conditions as strings
                string[] settings = config.Replace.Split("~");
                //Loop through splits
                foreach (string setting in settings)
                {
                    if (setting != null && setting != "")
                    {
                        //Split the conditions
                        string[] subSetting = setting.Split("|");
                        if (subSetting != null && subSetting.Count() > 0)
                        {
                            //Switch according to selection
                            cell_value = cell_value.Replace(subSetting[0].ToString(), subSetting[1].ToString());
                        }
                    }
                }
            }
            //End: Replace

            //Start: Extract
            if (config.Extract != "no")
            {
                switch (config.Extract)
                {
                    case "num":
                        cell_value = Regex.Replace(cell_value, "[^0-9]", "");
                        break;
                    case "txt":
                        cell_value = ExtractText(cell_value);
                        break;
                }
            }
            //End: Extract

            return cell_value;
        }

        //Extract only text from a string
        public static string ExtractText(string input)
        {
            // create a new StringBuilder to hold the extracted text
            StringBuilder sb = new StringBuilder();

            // loop through each character in the input string
            foreach (char c in input)
            {
                // if the character is a letter, append it to the StringBuilder
                if (Char.IsLetter(c))
                {
                    sb.Append(c);
                }
            }

            // convert the StringBuilder to a string and return it
            return sb.ToString();
        }
        #endregion

        #region: step - unpivot sub columns
        //Un Pivot data model
        private List<LabdipChartModel> arrangeSubColumns(List<LabdipChartModel> labdipChartModelList)
        {
            logger.InfoFormat("arrangeSubColumns function called with labdipChartModelList={0}", labdipChartModelList);

            //Initlaize variables
            List<LabdipChartModel> labdipChartModelListOutput = new List<LabdipChartModel>();

            try
            {
                //Check for model validity
                if (labdipChartModelList != null && labdipChartModelList.Count > 0)
                {
                    //Loop through row by row
                    foreach (LabdipChartModel labdipChartModel in labdipChartModelList)
                    {
                        //Check for exsistence of Sub Columns
                        if (labdipChartModel.SubColumns != null && labdipChartModel.SubColumns.Count > 0)
                        {
                            //Init row set for one row
                            List<LabdipChartModel> labDipChartNewRowSet = new List<LabdipChartModel>();
                            //Loop through Sub Columns
                            foreach (LabdipChartSubModel labdipChartSubModel in labdipChartModel.SubColumns)
                            {
                                //Skip Init rows
                                if (labdipChartSubModel.ColumnHeader != "INIT")
                                {
                                    // Get the PropertyInfo objects for the Id and field properties.
                                    PropertyInfo idProperty = typeof(LabdipChartModel).GetProperty(labdipChartSubModel.ColumnAttribute);
                                    PropertyInfo fieldProperty = typeof(LabdipChartModel).GetProperty(labdipChartSubModel.ValueAttribute);

                                    // Find the data model object with the specified Id value.
                                    LabdipChartModel model = labDipChartNewRowSet.FirstOrDefault(m => idProperty.GetValue(m).Equals(labdipChartSubModel.ColumnHeader));

                                    //Check is that any row exsist in Row Set
                                    if (model != null)
                                    {
                                        // Update the field value of the data model object.
                                        fieldProperty.SetValue(model, labdipChartSubModel.ColumnValue);
                                    }
                                    //Else add a new row to row set
                                    else
                                    {
                                        //Init and update exsisting values to new row
                                        LabdipChartModel labDipChartNewRow = new LabdipChartModel();
                                        labDipChartNewRow = UpdateModelObject(labdipChartModel);

                                        //Update fields
                                        labDipChartNewRow.GetType().GetProperty(labdipChartSubModel.ColumnAttribute).SetValue(labDipChartNewRow, labdipChartSubModel.ColumnHeader);
                                        labDipChartNewRow.GetType().GetProperty(labdipChartSubModel.ValueAttribute).SetValue(labDipChartNewRow, labdipChartSubModel.ColumnValue);
                                        //Add to row set
                                        labDipChartNewRowSet.Add(labDipChartNewRow);

                                        //Update flag
                                        labDipChartNewRow.Split = "sub";

                                    }
                                }
                            }
                            //Add rowset to main model list
                            labdipChartModelListOutput.AddRange(labDipChartNewRowSet);
                        }
                    }
                }
                //Output
                return labdipChartModelListOutput;
            }
            catch (Exception error)
            {
                logger.InfoFormat("error occured in arrangeSubColumns function called with labdipChartModelList={0}, error={1}", labdipChartModelList, error);
                _error = _error + error.Message + ". ";
                return null;
            }
        }

        //Update and set data LabdipChartModel
        public LabdipChartModel UpdateModelObject(LabdipChartModel labdip_chart)
        {
            LabdipChartModel labdip_chart_output = new LabdipChartModel();
            if (labdip_chart != null)
            {
                labdip_chart_output.Index = labdip_chart.Index;
                labdip_chart_output.RowIndex = labdip_chart.RowIndex;
                labdip_chart_output.Division = labdip_chart.Division;
                labdip_chart_output.Season = labdip_chart.Season;
                labdip_chart_output.Category = labdip_chart.Category;
                labdip_chart_output.Program = labdip_chart.Program;
                labdip_chart_output.StyleNoIndividual = labdip_chart.StyleNoIndividual;
                labdip_chart_output.GMTDescription = labdip_chart.GMTDescription;
                labdip_chart_output.GMTColor = labdip_chart.GMTColor;
                labdip_chart_output.NRF = labdip_chart.NRF;
                labdip_chart_output.ColorCode = labdip_chart.ColorCode;
                labdip_chart_output.RMColor = labdip_chart.RMColor;
                labdip_chart_output.PackCombination = labdip_chart.PackCombination;
                labdip_chart_output.PalcementName = labdip_chart.PalcementName;
                labdip_chart_output.BOMSelection = labdip_chart.BOMSelection;
                labdip_chart_output.ItemName = labdip_chart.ItemName;
                labdip_chart_output.SupplierName = labdip_chart.SupplierName;
                labdip_chart_output.RMColorRef = labdip_chart.RMColorRef;
                labdip_chart_output.GarmentWay = labdip_chart.GarmentWay;
                labdip_chart_output.FBNumber = labdip_chart.FBNumber;
                labdip_chart_output.MaterialType = labdip_chart.MaterialType;
                labdip_chart_output.ColorDyeingTechnic = labdip_chart.ColorDyeingTechnic;
                labdip_chart_output.SubColumns = labdip_chart.SubColumns;
                labdip_chart_output.error = labdip_chart.error;
            }

            return labdip_chart_output;
        }

        #endregion

        #region: step - split and transform data
        private List<LabdipChartModel> splitTransformData(List<LabdipChartModel> labdipChartModelData, int file_variation)
        {
            logger.InfoFormat("splitTransformData function called with labdipChartModel={0}, file_variation={1}", labdipChartModelData, file_variation);

            //Initlaize variables
            HeaderMasterService headerMasterService = new HeaderMasterService();
            DataSplitTransformationService dataSplitTransformationService = new DataSplitTransformationService();

            List<LabdipChartModel> LabdipChartModelOutput = new List<LabdipChartModel>();

            List<HeaderMaster> _headers = headerMasterService.GetHeaderList();
            if (_headers != null) { _headers = _headers.Where(c => c.TransformData != "no" && c.Variation == file_variation).ToList(); }

            if (_headers != null && _headers.Count > 0)
            {
                foreach (LabdipChartModel labdipChartModel in labdipChartModelData)
                {
                    foreach (HeaderMaster hedaer in _headers)
                    {
                        var propertyInfo = labdipChartModel.GetType().GetProperty((hedaer.HeaderType == "static" ? hedaer.HeaderAttribute : hedaer.SubValueAttribute), BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);
                        if (propertyInfo != null)
                        {
                            string colValue = (string)propertyInfo.GetValue(labdipChartModel);
                            if (colValue != null && colValue.ToString().Trim() != "" && colValue.ToString().Trim() != "Null")
                            {
                                if (hedaer.TransformData == "list")
                                {
                                    List<DataSplitTransformation> dataSplitTransformations = dataSplitTransformationService.GetDataSplitTransformationListbyValue(colValue, file_variation);
                                    if (dataSplitTransformations != null && dataSplitTransformations.Count > 0)
                                    {
                                        foreach (DataSplitTransformation dataSplitTransformation in dataSplitTransformations)
                                        {
                                            LabdipChartModel labdipChartModelNew = new LabdipChartModel();
                                            labdipChartModelNew = UpdateModelObject(labdipChartModel);
                                            labdipChartModelNew.Split = hedaer.TransformData + colValue;
                                            propertyInfo.SetValue(labdipChartModelNew, Convert.ChangeType(processCellValue(dataSplitTransformation.TransformedData, hedaer), propertyInfo.PropertyType));

                                            LabdipChartModelOutput.Add(labdipChartModelNew);
                                        }
                                    }
                                    else
                                    {
                                        LabdipChartModelOutput.Add(labdipChartModel);
                                    }
                                }
                                else
                                {
                                    string[] splits = colValue.Split(hedaer.TransformData);
                                    if(splits != null)
                                    {
                                        int count = 1, TransformationId = dataSplitTransformationService.GetMaxTranformationId();

                                        foreach(string split in splits)
                                        {
                                            LabdipChartModel labdipChartModelNew = new LabdipChartModel();
                                            labdipChartModelNew = UpdateModelObject(labdipChartModel);
                                            labdipChartModelNew.Split = hedaer.TransformData + colValue;
                                            propertyInfo.SetValue(labdipChartModelNew, Convert.ChangeType(processCellValue(split, hedaer), propertyInfo.PropertyType));

                                            LabdipChartModelOutput.Add(labdipChartModelNew);

                                            if(hedaer.SaveTransformVariation && split != colValue)
                                            {
                                                DataSplitTransformation newVariation = new DataSplitTransformation();
                                                newVariation.Id = TransformationId;
                                                newVariation.HeaderAttribute = (hedaer.HeaderType == "static" ? hedaer.HeaderAttribute : hedaer.SubValueAttribute);
                                                newVariation.SubId = count;
                                                newVariation.Variation = file_variation;
                                                newVariation.InitialData = colValue;
                                                newVariation.TransformedData = split;

                                                dataSplitTransformationService.InsertNewRecord(newVariation);
                                            }

                                            count++;
                                        }
                                    }
                                    else
                                    {
                                        LabdipChartModelOutput.Add(labdipChartModel);
                                    }
                                }
                            }
                            else
                            {
                                LabdipChartModelOutput.Add(labdipChartModel);
                            }
                        }
                        else
                        {
                            LabdipChartModelOutput.Add(labdipChartModel);
                        }
                    }
                }
            }
            else
            {
                LabdipChartModelOutput = labdipChartModelData;
            }

            return LabdipChartModelOutput;
        }

        //Get a specific value from the model
        public static T GetValueFromModel<T>(object model, string fieldName)
        {
            var type = model.GetType();
            var propInfo = type.GetProperty(fieldName);
            var value = (T)propInfo.GetValue(model);
            return value;
        }

        #endregion
        #endregion
        #endregion
    }
}
