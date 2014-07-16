using System;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using System.Linq;
using MyExcelExport;

namespace MySample
{
    delegate void OnCellData(char Column, int RowNumber, string value);                                             // Callback that will process our data. Warning MAX NUMBER OF COLUMNS is 25. From 'A' to  'Z'

    class ExcelReader
    {
        OoXml source;
        gSheet sheet;
        bool IsText;
        char    NumCol;
        int     NumRow;
        OnCellData OnData;

        public ExcelReader(string filename)
        {
            source = new OoXml(filename);                                                                           // Open the excel file
            sheet = source.sheets.Values.First<gSheet>();                                                           // Get the first sheet
        }

        public void Process(OnCellData ondata)
        {
            OnData = ondata;                                                                                        // This is the callback that will receive the data
            Stream inStream = sheet.GetStream();                                                                    // This is the data stream
            using (XmlReader reader = XmlReader.Create(inStream))                                                   // We will read the stream as XML
            LoopSheet(reader);                                                                                      // So lets do it
        }
        
        private void GetCellInfo(XmlReader r)
        {
            IsText = false;                                                                                         // Let's assume this is a numeric cell
            NumCol = ' ';                                                                                           // With no column
            while (r.MoveToNextAttribute())                                                                         // Loop all attributes
            {   
                if (r.Name == "t") if (r.Value == "s") IsText = true;                                               // If there is a [t="s"] attribute  then this is a TEXT cell   
                if (r.Name == "r") if (r.Value.Length > 1) NumCol = r.Value[0];                                     // If there is a [r="Cnn"] attribute get column from C
            }
        }

        private void LoopSheet(XmlReader r)
        {
            bool InValue = false;                                                                                   // Flag that tells if we are in a data node    
            while (r.Read())                                                                                        // While reading nodesdata
                switch (r.NodeType)                                                                                 // If the node is
                {
                    case XmlNodeType.Element:                                                                       // An open element
                        InValue = false;                                                                            //      Lets assume is a NON data node
                        if (r.Name == "row") { OnData('-', NumRow, ""); NumRow++; break; }                          //      IF this is a ROW node, tell the host app and increase NumRos
                        if (r.Name == "c") { GetCellInfo(r); break; }                                               //      If this is a CELL node, get cell info
                        if (r.Name == "v") InValue = true;                                                          //      IF it turns out that this is a data node
                        break;
                    case XmlNodeType.EndElement:                                                                    // A close element        
                        InValue = false;                                                                            //      For sure this is not a data node
                        if (r.Name == "row") OnData('#', NumRow, "");                                               //      IF a row has completed, warn the host app
                        break;
                    case XmlNodeType.Text:                                                                          // Data content
                        if (!InValue) break;                                                                        //      Skip if not i a data node
                        string s = r.Value;                                                                         //      Get data (the parser always returns a string)
                        s = (IsText ? source.words[Int32.Parse(s)] : s);                                            //      If the string points to a text reference, translate the text reference
                        OnData(NumCol, NumRow,s);                                                                   //      Inform the host the actual readed data 
                        break;
                }
        }
    }
}
