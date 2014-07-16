using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using MyExcelExport;

namespace MySample
{
    class ExcelWriter
    {
        OoXml                   template;
        gSheet                  sheet;
        IDataReader             data;
        Dictionary<int, string> mapping = new Dictionary<int, string>() ;
        int                     fixedRows = 0;
        int                     curRow = 1;
        int                     exportedRecords = 0;

        public int NumRecords { get { return exportedRecords; } }
        
        public ExcelWriter(string filename)
        {
            template = new OoXml(filename);                                         // Open the template
            sheet = template.sheets.Values.First<gSheet>();                         // Get the first sheet
            GetMapping();                                                           // Get the mapping    
            sheet.SetSource(OnDataRow, fixedRows);                                                                          // Set data call back, keep first row of the template         
        }

        public void Export(IDataReader dataSource, string filename)
        {
            data = dataSource;                                                      // This is the data source             
            curRow    = fixedRows+1;                                                // The first row written will be after the FixedRows                                                  
            exportedRecords = 0;                                                    // The number of rows we have exported
            template.Save(filename);
        }
        
        void OnDataRow(gSheet sheet)
        {
            if (!data.Read()) return;                                               // If no data available then exit
            sheet.BeginRow(curRow);                                                 // This is the row we are going to write
            foreach(KeyValuePair<int,string> m in mapping)                          // For each mapping value
                sheet.WriteCell(m.Key, data[m.Value]);                              // Save it into the Excel
            sheet.EndRow();                                                         // Data has been exported
            curRow++;                                                               // Next record will be placed in this excel row
            exportedRecords++;                                                      // Increase number of exported records
        }

        private void GetMapping()
        {
            string[] cells;                                                         // Cells of current row
            int i = 0;                                                              // First row to review
            do                                                                      // Loop all the rows
            {                                                                               
                cells = sheet.Row(i);                                               // Get cells for row [i]. WARNING template cannot be a BIG file. It's parsed fully in memory    
                if (cells != null) CheckMapping(i, cells);                          // If there are cells, check if they define a mapping
                i++;                                                                // Next row
            } while (cells != null && mapping.Count == 0);                          // Exit when there are no rows or a mapping has been found 
        }

        private void CheckMapping(int row, string[] cells)                          // Checks if a row defines a mapping
        {
            for (int i = 0; i < cells.Length; i++)                                  // Loop all the cells in the reow
                if (cells[i].Left(1) == "*")                                        // if a cell starts with "*"
                    mapping[i] = cells[i].Replace("*", "");                         // The it defines a mapping, save it        
            if (mapping.Count > 0) fixedRows = row;                                 // Rows below the mapping will be preserved    
        }
    }
}
