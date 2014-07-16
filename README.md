

Just a bunch of C# (.Net 3.5+) classes for REALLY FAST exporting & importing of Excel OoXMl (2007+) files.

The library is totally self-contained, it does not require any other piece software is required (no Excel needed, no ODBC, ACE OLEDB, OpenXML ,....)

Although this is a .Net library and memory control is rather limited, memory consumption is kept as low as possible, and it doesn't depend on the number of rows being transferred.

Included is a sample that:

Open a XLSX template
Fills the template with 100,000 rows using a IDataReader as a source. Column ordered is decided by matching the template contents with the data source field names.
Saves the template with another name.
Opens the template
Reads 100,000 rows from the template
