using System;
using System.Collections.Generic;
using System.Data;
using System.Text;


namespace MySample
{

    class Program
    {

        static void Main(string[] args)
        {
            string output_file="Completed_Template.xlsx";
            DateTime t0 = DateTime.Now;
            ExcelWriter t = new ExcelWriter("Template.xlsx");
            MySampleData data = new MySampleData();
            t.Export(data,output_file);
            DateTime t1=DateTime.Now;
            Console.WriteLine("Written " + t.NumRecords + " records to file ["+output_file+"] in "+(t1-t0).TotalSeconds+" seconds");
            ExcelReader r = new ExcelReader(output_file);
            r.Process(OnExcelCell);
            DateTime t2=DateTime.Now;
            Console.WriteLine("Readed   " + t.NumRecords + " rows from file [" + output_file + "] in " + (t2 - t1).TotalSeconds + " seconds");
        }

        // This function gets calledfor each cell readed from excel
        static void OnExcelCell(char Column, int RowNumber, string value)
        {
            /* Uncomment this block to see what is written
             
            if (RowNumber <=5)
            {
                if (Column == '-') Console.Write("     " + RowNumber.ToString().PadRight(5));
                else if (Column == '#') Console.WriteLine("");
                else
                    Console.Write(Column + ": " + value.PadRight(40).Substring(0, 40) + "  ");
                return;
            }
            if (RowNumber == 6 && Column == '#')
            {
                Console.WriteLine("     rows over 10 are not displayed :)");
                Console.WriteLine(" ");
            }
             * */
        }

    }
}
