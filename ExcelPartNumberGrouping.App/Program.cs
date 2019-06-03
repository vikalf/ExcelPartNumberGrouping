using Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelPartNumberGrouping.App
{
    class Program
    {
        public static readonly string BOM_MASTER_FILE_LOCATION = @"D:\Jirux\MasterBOM.xlsx";
        public static readonly string FAMILIES_TO_FILTER = "720618,720619,720624,720626,720627,720630,71000037,71000653,71000654,720631,720632,720633,71000038,720616,,71000036,71000045,720606,720609,71000132,720607,71000023,51000075,51000079,71000024,51000076,51000080,71000025,51000077,51000081,71000026,51000078,51000082,71000044,720644,7101120,720645,71000075,71000066,71000068,71000169,71000134,71000137,71000051,71000052,71000057,71000058,71000053,7101106,71000063,71000140,71000141,71000142,71000143,71000651,7101058,7101054,7101110,7101112,7101161,71000144,7101076,7101132,7101132,71000643,7101055,7101057,7101070,7101083,7101084,7101148,7101011,7101072,7101128,7101093,7101098,7101096,7101101,7101146,7101145,7101188,7101191,7101190,7101187,7101102,7101103,7101108,7101111,5101473,7101147,7101223,7101125,7101126,7101127,7101137,7101138,7101150,7101151,7101129,7101131,7101163,7101174,7101164,7101196";
        public static readonly string OUTPUT_FILE_PATH = @"D:\temp";

        static void Main(string[] args)
        {

            try
            {
                // 1.- Get Data from Spreadsheet
                List<PartNumberDetailedDao> partNumberDetails = GetPartNumberDetailsDao();

                var partNumbers = GetPartNumbers(partNumberDetails);

                var families = GetFamilies(partNumberDetails);


                // 4.- Now I have the data.. Let's Generate Excel File!!!
                GenerateExcelFile(partNumbers, families);

                Console.WriteLine("Completed..");

                Console.ReadKey();
            }
            catch (Exception error)
            {
                // OOPS! an error occurred, what happened??
                Console.WriteLine(error.ToString());
            }
        }

        private static void GenerateExcelFile(List<PartNumberItem> partNumbers, List<Family> families)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {

                excel.Workbook.Worksheets.Add("ProductionPlan (MASTER)");

                var worksheet = excel.Workbook.Worksheets["ProductionPlan (MASTER)"];
                worksheet.Cells[2, 1].Value = "Part Number";
                worksheet.Cells[2, 2].Value = "Description";

                int partNumberRow = 3;
                // Fill All Part Numbers 
                foreach (var partNumber in partNumbers)
                {
                    worksheet.Cells[partNumberRow, 1].Value = partNumber.ItemNumber;
                    worksheet.Cells[partNumberRow, 2].Value = partNumber.Description;
                    partNumber.FileRow = partNumberRow;
                    partNumberRow++;
                }

                int familyNumberColumn = 3;
                // Fill Families
                foreach (var family in families)
                {
                    worksheet.Cells[2, familyNumberColumn].Value = family.ItemNumber;
                    family.FileColumn = familyNumberColumn;

                    foreach (var partNumber in family.PartNumbers)
                    {
                        var partNumberItem = partNumbers.FirstOrDefault(e => e.ItemNumber.Trim() == partNumber.ItemNumber.Trim());

                        if (partNumber != null)
                            worksheet.Cells[partNumberItem.FileRow, familyNumberColumn].Value = partNumber.Quantity.ToString();

                    }

                    familyNumberColumn++;
                }


                FileInfo excelFile = new FileInfo($"{OUTPUT_FILE_PATH}\\Output_Jirux.xlsx");
                excel.SaveAs(excelFile);

            }
        }

        private static List<Family> GetFamilies(List<PartNumberDetailedDao> partNumberDetails)
        {
            var familiesToFilter = FAMILIES_TO_FILTER.Split(',').Select(e => e.Trim()).ToList();

            var filteredFamilies = partNumberDetails
                .Where(e => !string.IsNullOrEmpty(e.ItemNumberParent.Trim()) &&
                familiesToFilter.Contains(e.ItemNumberParent.Trim()))
                .Select(e => e.ItemNumberParent).Distinct().ToList();

            var families = filteredFamilies.Select(e => new Family
            {
                ItemNumber = e.Trim(),
                Description = partNumberDetails.FirstOrDefault(f => f.ItemNumberParent.Trim() == e)?.DescriptionParent ?? string.Empty,
                PartNumbers = new List<PartNumber>()
            }).ToList();

            foreach (var family in families)
            {
                var partNumbers = partNumberDetails.Where(e => e.ItemNumberParent.Trim() == family.ItemNumber);

                family.PartNumbers = partNumbers.Select(e => new PartNumber
                {
                    Description = e.DescriptionChild,
                    ItemNumber = e.ItemNumberChild,
                    Level = e.Level,
                    Quantity = e.Quantity,
                    UM = e.UM
                }).ToList();

            }

            /// TEST ///


            var testFamily = families.FirstOrDefault(e => e.ItemNumber == "71010068");

            var cleanedUpPartNumbers = CleanUpPartNumbersLogic(testFamily.PartNumbers, testFamily.Description);

            /////

            return families;
        }

        private static List<PartNumber> CleanUpPartNumbersLogic(List<PartNumber> partNumbers, string familyDescription)
        {

            if (partNumbers == null || !partNumbers.Any())
                return new List<PartNumber>();

            var result = partNumbers;

            
            var totalPartNumbers = result.Count;
            var kfgEvaluated = false;
            for (int i = 0; i < result.Count; i++)
            {
                kfgEvaluated = false;
                // Check "KGF"
                if (result[i].Description.Trim().StartsWith("KGF"))
                {
                    result[i].BOMIncluded = true;
                    var kgfIndex = i;
                    var KgfPartNumberLevel = int.Parse(result[i].Level[result[i].Level.Length - 1].ToString());

                    for (int j = kgfIndex; j < totalPartNumbers; j++)
                    {
                        var nextKgfPartNumber = result[j + 1];
                        if (nextKgfPartNumber != null)
                        {
                            var nextKgfPartNumberLevel = int.Parse(nextKgfPartNumber.Level[nextKgfPartNumber.Level.Length - 1].ToString());

                            if (nextKgfPartNumberLevel <= KgfPartNumberLevel)
                            {
                                i = j;
                                kfgEvaluated = true;
                                break;
                            }

                        }
                        else
                            break;

                    }
                }

                if (kfgEvaluated)
                {
                    result[i + 1].BOMIncluded = true;
                    continue;
                }
                    

                if (i >= result.Count - 1)
                {
                    result[i].BOMIncluded = true;
                    break;
                }
                    
                


                var nextPartNumber = result[i + 1];
                var currentPartNumberLevel = int.Parse(result[i].Level[result[i].Level.Length - 1].ToString());

                if (nextPartNumber != null)
                {
                    

                    var nextPartNumberLevel = int.Parse(nextPartNumber.Level[nextPartNumber.Level.Length - 1].ToString());

                    if (nextPartNumberLevel <= currentPartNumberLevel)
                        result[i].BOMIncluded = true;
                    else
                    {
                        if (result[i].UM.Trim() != nextPartNumber.UM.Trim() || nextPartNumber.Description.ToLowerInvariant().Contains("terminal"))
                            result[i].BOMIncluded = true;
                    }
                }


            }

            

            return result.Where(e => e.BOMIncluded).ToList();

        }

        private static List<PartNumberItem> GetPartNumbers(List<PartNumberDetailedDao> partNumberDetails)
        {
            var filteredPartNumbers = partNumberDetails
                .Where(e => !string.IsNullOrEmpty(e.ItemNumberChild.Trim()))
                .Select(e => e.ItemNumberChild).Distinct();

            var result = filteredPartNumbers.Select(e => new PartNumberItem
            {
                Description = partNumberDetails.FirstOrDefault(f => f.ItemNumberChild.Trim() == e)?.DescriptionChild ?? string.Empty,
                ItemNumber = e
            }).ToList();

            return result;
        }

        /// <summary>
        /// 1.- Get Part Number Detail 
        /// </summary>
        /// <returns></returns>
        private static List<PartNumberDetailedDao> GetPartNumberDetailsDao()
        {
            var excelfileExtension = Path.GetExtension(BOM_MASTER_FILE_LOCATION);
            using (FileStream stream = File.Open(BOM_MASTER_FILE_LOCATION, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader excelReader = getExcelReader(excelfileExtension, stream))
                {
                    excelReader.IsFirstRowAsColumnNames = true;

                    DataSet familyDs = excelReader.AsDataSet();
                    DataTable familyDT = familyDs.Tables[0];

                    List<PartNumberDetailedDao> partNumberDetailsDAO = new List<PartNumberDetailedDao>();

                    foreach (DataRow row in familyDT.Rows)
                        partNumberDetailsDAO.Add(MapPartNumberDetailedDAO(row));

                    return partNumberDetailsDAO;
                }

            }
        }



        #region Helpers

        private static PartNumberDetailedDao MapPartNumberDetailedDAO(DataRow row)
        {
            return new PartNumberDetailedDao
            {
                Level = row[0].ToString(),
                ItemNumberParent = row[1].ToString(),
                DescriptionParent = row[2].ToString(),
                ItemNumberChild = row[3].ToString(),
                DescriptionChild = row[4].ToString(),
                Quantity = int.Parse(row[5].ToString()),
                UM = row[6].ToString(),
            };
        }

        private static IExcelDataReader getExcelReader(string excelfileExtension, FileStream stream)
        {
            IExcelDataReader excelReader;

            if (excelfileExtension == ".xls" || excelfileExtension == ".xlsb")
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            else
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            return excelReader;
        }

        #endregion

        #region Entities

        public class PartNumberDetailedDao
        {
            public string Level { get; set; }
            public string DescriptionParent { get; set; }
            public string ItemNumberParent { get; set; }
            public string DescriptionChild { get; set; }
            public string ItemNumberChild { get; set; }
            public int Quantity { get; set; }
            public string UM { get; set; }
        }

        public class Family
        {
            public string Description { get; set; }
            public string ItemNumber { get; set; }
            public List<PartNumber> PartNumbers { get; set; }
            public int FileColumn { get; set; }

        }

        public class PartNumberItem
        {
            public string Description { get; set; }
            public string ItemNumber { get; set; }
            public int FileRow { get; set; }
        }

        public class PartNumber
        {
            public string Level { get; set; }
            public string Description { get; set; }
            public string ItemNumber { get; set; }
            public int Quantity { get; set; }
            public string UM { get; set; }
            public bool BOMIncluded { get; set; }

        }

        #endregion




    }
}
