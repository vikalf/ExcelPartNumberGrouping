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
        public static readonly string PART_NUMBERS_FILE_LOCATION = @"C:\temp\PartNumbers.xlsx";
        public static readonly string FAMILY_PART_NUMBERS_FILE_LOCATION = @"C:\temp\FamilyPartNumbers.xlsx";
        public static readonly string OUTPUT_FILE_PATH = @"C:\temp";

        static void Main(string[] args)
        {
            try
            {
                // 1.- Get Families
                List<FamilyPartNumber> families = GetFamilyPartNumbers();

                // 2.- Get Part Numbers
                List<PartNumber> partNumbers = GetPartNumbers();


                // 3 .- Get Group Part Numbers by Family
                List<FamilyPartNumberGrouped> familyPartNumberGroupeds = GroupFamilyPartNumbers(families, partNumbers);

                // 4.- Now I have the data.. Let's Generate Excel File!!!
                GenerateExcelFile(familyPartNumberGroupeds);

                Console.WriteLine("Completed..");
                Console.ReadKey();
            }
            catch (Exception error)
            {
                // OOPS! an error occurred, what happened??
                Console.WriteLine(error.ToString());
            }
        }

        /// <summary>
        /// 1.- Get Families
        /// </summary>
        /// <returns></returns>
        private static List<FamilyPartNumber> GetFamilyPartNumbers()
        {
            var excelfileExtension = Path.GetExtension(FAMILY_PART_NUMBERS_FILE_LOCATION);
            using (FileStream stream = File.Open(FAMILY_PART_NUMBERS_FILE_LOCATION, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader excelReader = getExcelReader(excelfileExtension, stream))
                {
                    excelReader.IsFirstRowAsColumnNames = true;

                    DataSet familyDs = excelReader.AsDataSet();
                    DataTable familyDT = familyDs.Tables[0];

                    List<FamilyPartNumber> familyParNumbersDAO = new List<FamilyPartNumber>();

                    foreach (DataRow row in familyDT.Rows)
                        familyParNumbersDAO.Add(MapFamilyPartNumbersDAO(row));

                    return familyParNumbersDAO;
                }

            }
        }

        /// <summary>
        /// 2.- Get Part Numbers
        /// </summary>
        private static List<PartNumber> GetPartNumbers()
        {
            var excelfileExtension = Path.GetExtension(PART_NUMBERS_FILE_LOCATION);
            using (FileStream stream = File.Open(PART_NUMBERS_FILE_LOCATION, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader excelReader = getExcelReader(excelfileExtension, stream))
                {
                    excelReader.IsFirstRowAsColumnNames = true;

                    DataSet partNumbersDs = excelReader.AsDataSet();
                    DataTable partNumbersDT = partNumbersDs.Tables[0];

                    List<PartNumber> parNumbersDAO = new List<PartNumber>();

                    foreach (DataRow row in partNumbersDT.Rows)
                        parNumbersDAO.Add(MapPartNumbersDAO(row));

                    return parNumbersDAO;
                }

            }

        }

        /// <summary>
        /// 3 .- Get Group Part Numbers by Family
        /// </summary>
        private static List<FamilyPartNumberGrouped> GroupFamilyPartNumbers(List<FamilyPartNumber> families, List<PartNumber> partNumbers)
        {
            List<FamilyPartNumberGrouped> result = new List<FamilyPartNumberGrouped>();

            foreach (var family in families)
            {

                result.Add(new FamilyPartNumberGrouped
                {
                    Family = family,
                    PartNumbers = partNumbers.Where(e => e.FamilyPartID.Trim() == family.ID.Trim()).ToList()
                });
            }
            return result;
        }

        /// <summary>
        /// 4.- Now I have the data.. Let's Generate Excel File!!!
        /// </summary>
        private static void GenerateExcelFile(List<FamilyPartNumberGrouped> familyPartNumberGroupeds)
        {
            int offset = 0;
            var chunkSize = 500;
            var chunkCount = Math.Ceiling(familyPartNumberGroupeds.Count / (decimal)chunkSize);


            for (int i = 0; i < chunkCount; i++)
            {
                var familiesToTake = Math.Min(chunkSize, familyPartNumberGroupeds.Count - i * chunkSize);
                var chunk = familyPartNumberGroupeds.Skip(offset).Take(familiesToTake).ToArray();

                using (ExcelPackage excel = new ExcelPackage())
                {
                    foreach (var familyGrouped in chunk)
                    {
                        excel.Workbook.Worksheets.Add(familyGrouped.Family.ID);

                        var worksheet = excel.Workbook.Worksheets[familyGrouped.Family.ID];

                        worksheet.Cells[1, 1].Value = "Familia";
                        worksheet.Cells[1, 2].Value = "Familia Descripcion";
                        worksheet.Cells[1, 3].Value = "No. Parte";
                        worksheet.Cells[1, 4].Value = "No. Parte Description";

                        var index = 2;

                        foreach (var childPartNode in familyGrouped.PartNumbers)
                        {
                            worksheet.Cells[index, 1].Value = familyGrouped.Family.ID;
                            worksheet.Cells[index, 2].Value = familyGrouped.Family.Description;
                            worksheet.Cells[index, 3].Value = childPartNode.ID;
                            worksheet.Cells[index, 4].Value = childPartNode.Description;
                            index++;
                        }
                    }

                    FileInfo excelFile = new FileInfo($"{OUTPUT_FILE_PATH}\\Output_{i}.xlsx");
                    excel.SaveAs(excelFile);

                }

                offset += (chunk.Length);
            }



            

        }

        #region Helpers

        private static PartNumber MapPartNumbersDAO(DataRow row)
        {

            return new PartNumber
            {
                ID = row[0].ToString(),
                Description = row[1].ToString(),
                FamilyPartID = row[2].ToString(),
            };
        }

        private static FamilyPartNumber MapFamilyPartNumbersDAO(DataRow row)
        {
            return new FamilyPartNumber
            {
                ID = row[0].ToString(),
                Description = row[1].ToString(),
            };
        }

        private static IExcelDataReader getExcelReader(string excelfileExtension, FileStream stream)
        {
            IExcelDataReader excelReader;

            if (excelfileExtension == ".xls")
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            else
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            return excelReader;
        }

        #endregion

        #region Entities

        public class FamilyPartNumber
        {
            public string ID { get; set; }
            public string Description { get; set; }
        }

        public class FamilyPartNumberGrouped
        {
            public FamilyPartNumber Family { get; set; }
            public List<PartNumber> PartNumbers { get; set; }
        }

        public class PartNumber
        {
            public string ID { get; set; }
            public string Description { get; set; }
            public string FamilyPartID { get; set; }
        }

        #endregion




    }
}
