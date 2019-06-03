using System;

namespace BOMMaster.Appp
{

    class Program
    {
        public static readonly string BOM_MASTER_FILE_LOCATION = @"D:\Jirux\MasterBOM.xlsb";

        static void Main(string[] args)
        {



            Console.WriteLine("Hello World!");
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
