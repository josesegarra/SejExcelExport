using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Xml;

namespace MyExcelExport
{
    
    static class Extension
    {
        public static string Attribute(this XmlNode n, string name)
        {
            XmlAttribute b = n.Attributes[name];
            if (b == null) return ""; else return b.Value;
        }
    }


    public delegate void OnNewRow(gSheet sheet);
    
    // http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.cell(v=office.14).aspx
    class Column
    {
        internal byte[] part1, part2, part3;
        internal Column(XmlNode n)
        {
            string p="      <c";                                                                                                // This is the cell name    
            foreach(XmlAttribute a in n.Attributes) if (a.Name!="r" && a.Name!="t") p=p+" "+a.Name+"=\""+a.Value+"\"";          // Get all attributes except r and t
            string ci=new String(n.Attribute("r").Where(c => (c > '9')).ToArray());                                              // Get col Index as [r] without numeric chars
            part1=ASCIIEncoding.ASCII.GetBytes(p+" r=\""+ci);
            part2=ASCIIEncoding.ASCII.GetBytes("\"><v>");
            part3=ASCIIEncoding.ASCII.GetBytes("</v></c>\n");
        }
    }


    // This Stream implements Writing Rows of eBilling Data To Sheet Stream in Excel file.
    // On creation it creates Byte[] holding the common part of the received sheet
    class OoXmlDataStream : MemoryStream
    {
        List<Column> columns = new List<Column>();                                                                              // Format of columns to be replaces
        byte[]      header;                                                                                                     // Xml before data
        byte[]      footer;                                                                                                     // Xml after data
        byte[]      rowopen1 ,rowopen2,rowclose;                                                                                // Call back on writing row
        OnNewRow    onrow;
        gSheet      srsheet;
        bool        datadone = false;
        
        int         nbytes;
        byte[]      nbuffer;
        byte[]      thisrow;
        byte[]      sstpart;
        public OoXmlDataStream(gSheet sheet, int Keep,OnNewRow dr)
            : base()
        {
            string separator = Guid.NewGuid().ToString();
            sheet.LoadInMemory();                                                                                               // Make sure that the sheet is loaded in memory
            XmlDocument d = sheet.Stream.ReadAsXml();                                                                           // Load as XML
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(d.NameTable);                                                   // Set the name space    
            nsmgr.AddNamespace("aa", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");                              // To Open XML
            XmlNode ws = d.SelectSingleNode("//aa:worksheet/aa:sheetData", nsmgr);                                              // Locate data NODE
            if (ws == null) throw new Exception("Could not find sheet data");                                                   // Error if not
            XmlNode r = ws.FirstChild; while (Keep > 0 && r != null) { r = r.NextSibling; Keep--; }                             // Skip rows to keep
            if (r != null) InitializeColumns(r);                                                                                // if present use first row NOT to keep to initialize columns
            ws.AppendChild(d.CreateNode(XmlNodeType.Comment, "SEPARATOR", null)).InnerText = separator;                         // Add separator mark up
            InitializeHeaderAndFooter(d.OuterXml, separator);                                                                   // Get header and footer using separator
            rowopen1= ASCIIEncoding.ASCII.GetBytes("<row r=\"");                                                                // <row r="
            sstpart  = ASCIIEncoding.ASCII.GetBytes("\" t=\"s");                                                                // " t="s 
            rowopen2 = ASCIIEncoding.ASCII.GetBytes("\">\n");                                                                   // ">
            rowclose = ASCIIEncoding.ASCII.GetBytes("</row>\n");                                                                // </row>
          
            srsheet = sheet;                                                                                                    // Sheet this stream is linked to
            onrow = dr;                                                                                                         // On data callback
        }

        private void InitializeColumns(XmlNode r)
        {
            XmlNode c = r.FirstChild;                                                                                           // This is the first column
            while (c != null) { if (c.Name == "c") columns.Add(new Column(c)); c = c.NextSibling; }                             // Loop all columns an initialize their format
            while (r.NextSibling != null) r.ParentNode.RemoveChild(r.NextSibling);                                              // Remove all rows after r
            r.ParentNode.RemoveChild(r);                                                                                        // Remove r
        }

        private void InitializeHeaderAndFooter(string s, string r)                                                              // Initialize global header & footer
        {
            int i= s.IndexOf(r);                                                                                                // Use separator to split XML
            while (s[i] != '<') i--;header=ASCIIEncoding.ASCII.GetBytes(s.Substring(0, i));                                     // First part will be header
            while (s[i] != '>') i++;footer=ASCIIEncoding.ASCII.GetBytes(s.Substring(i+1));                                      // Second part will be footer            
        }

        public override int Read(byte[] buffer, int start, int count)                                                           // This one is called by the writer    
        {
            byte[] d;
            if (header != null) { d = header; header = null; return WriteBytes(d, buffer, 0); }                                 // If header has not been written. Write it and return
            if (!datadone)                                                                                                      // If all data has not been written    
            {
                nbytes = 0; nbuffer = buffer;                                                                                   // Reset the number of bytes written
                onrow(srsheet);                                                                                                 // Ask the host to write the data
                if (nbytes>0) return nbytes;                                                                                    // If host writted something then return it
            }
            datadone = true;                                                                                                    // If we ever get here, all data has already been written
            if (footer != null) { d = footer; footer = null; return WriteBytes(d, buffer, 0); }                                 // If footer has not been written. Write it and return
            return 0;
        }
        
        private int WriteBytes(byte[] w, byte[] buffer,int pos)
        {
            int j = w.Length;
            if (buffer.Length < j) throw new Exception("Partial buffer writing NOT implemented");
            System.Buffer.BlockCopy(w, 0, buffer, pos, j);
            return j;
        }

        private int WriteInt(int k, byte[] buffer,int pos)
        {
            byte[] w = ASCIIEncoding.ASCII.GetBytes(k.ToString());
            int j = w.Length;
            if (buffer.Length > j) { } else throw new Exception("Partial buffer writing NOT implemented");
            System.Buffer.BlockCopy(w, 0, buffer, pos, j);
            return j;
        }

        public void BeginRow(int RowNumber) 
        {
            nbytes = nbytes + WriteBytes(rowopen1, nbuffer, nbytes);
            thisrow= ASCIIEncoding.ASCII.GetBytes(RowNumber.ToString());
            nbytes = nbytes + WriteBytes(thisrow, nbuffer, nbytes); 
            nbytes = nbytes + WriteBytes(rowopen2, nbuffer, nbytes);                // Write row TAG

        }

        public void EndRow() 
        {
            nbytes = nbytes + WriteBytes(rowclose, nbuffer, nbytes);
        }

        public void WriteCell(int cell, int value)
        {
            if (cell >= columns.Count) return;
            Column c=columns[cell];
            nbytes = nbytes + WriteBytes(c.part1,nbuffer,nbytes);
            nbytes = nbytes + WriteBytes(thisrow, nbuffer, nbytes); 
            nbytes = nbytes + WriteBytes(c.part2,nbuffer,nbytes);
            nbytes = nbytes + WriteInt(value, nbuffer, nbytes); 
            nbytes = nbytes + WriteBytes(c.part3,nbuffer,nbytes);
        }

        public void WriteSSTCell(int cell, int value)
        {
            if (cell >= columns.Count) return;
            Column c = columns[cell];
            nbytes = nbytes + WriteBytes(c.part1, nbuffer, nbytes);
            nbytes = nbytes + WriteBytes(thisrow, nbuffer, nbytes);
            nbytes = nbytes + WriteBytes(sstpart, nbuffer, nbytes);
            nbytes = nbytes + WriteBytes(c.part2, nbuffer, nbytes);
            nbytes = nbytes + WriteInt(value, nbuffer, nbytes);
            nbytes = nbytes + WriteBytes(c.part3, nbuffer, nbytes);
        }
        
        public void WriteCell(int cell, double value) 
        {
            if (cell >= columns.Count) return;
            Column c = columns[cell];
            nbytes = nbytes + WriteBytes(c.part1, nbuffer, nbytes);
            nbytes = nbytes + WriteBytes(thisrow, nbuffer, nbytes);
            nbytes = nbytes + WriteBytes(c.part2, nbuffer, nbytes);
            byte[] w = ASCIIEncoding.ASCII.GetBytes(value.ToString());
            nbytes = nbytes + WriteBytes(w, nbuffer, nbytes);
            nbytes = nbytes + WriteBytes(c.part3, nbuffer, nbytes);
        }


    }
}
