using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOM.Master.App
{
    class Program
    {
        public static readonly string BOM_MASTER_FILE_LOCATION = @"D:\Jirux\MasterBOM.xlsb";

        static void Main(string[] args)
        {



            Console.WriteLine("Hello World!");
        }


        private static List<PartNumberDetailedDao> GetPartNumberDetailsDao()
        {
            var excelfileExtension = Path.GetExtension(BOM_MASTER_FILE_LOCATION);
            using (FileStream stream = File.Open(BOM_MASTER_FILE_LOCATION, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader excelReader = getExcelReader(excelfileExtension, stream))
                {
                    //excelReader.isfiIsFirstRowAsColumnNames = true;

                    DataSet familyDs = excelReader.AsDataSet();
                    DataTable familyDT = familyDs.Tables[0];

                    List<PartNumberDetailedDao> partNumberDetailsDAO = new List<PartNumberDetailedDao>();

                    foreach (DataRow row in familyDT.Rows)
                        partNumberDetailsDAO.Add(MapFamilyPartNumbersDAO(row));

                    return partNumberDetailsDAO;
                }

            }
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
        }

        public class PartNumber
        {
            public string Level { get; set; }
            public string Description { get; set; }
            public string ItemNumber { get; set; }
            public int Quantity { get; set; }
            public string UM { get; set; }

        }


    }
}
