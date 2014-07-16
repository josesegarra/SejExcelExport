using System;
using System.Data;


// Jose: This is just a sample object that implements IDataReader. 
// it returns a lot of records with dummy values in each column

namespace MySample
{
    class MySampleData : IDataReader
    {
        int currow = 0;
        int maxrow = 100000;
        Random rnd = new Random();

        public bool Read()
        {
            if (currow <= maxrow) currow++;
            return (currow <= maxrow);
        }

        public object this[String name]
        {
            get 
            { 
                if (name=="age") return rnd.Next(1, 90);
                return "Value of Column [" + name + "] for Row [" + currow + "]"; 
            }
        }

        // ############################# DUMMY methods so the IDataReader interface is implemented
        public void Close() { }
        public int Depth { get { return 0; } }
        public bool IsClosed { get { return false; } }
        public int RecordsAffected { get { return -1; } }
        public bool NextResult() { return false; }

        public DataTable GetSchemaTable() { throw new NotSupportedException(); }
        public String GetName(int i) { throw new NotSupportedException(); }

        public String GetDataTypeName(int i) { throw new NotSupportedException(); }
        public Type GetFieldType(int i) { throw new NotSupportedException(); }
        public int FieldCount { get { throw new NotSupportedException(); } }
        public Object GetValue(int i) { throw new NotSupportedException(); }
        public int GetValues(object[] values) { throw new NotSupportedException(); }
        public int GetOrdinal(string name) { throw new NotSupportedException(); }
        public object this[int i] { get { throw new NotSupportedException(); } }
        public bool GetBoolean(int i) { throw new NotSupportedException(); }
        public byte GetByte(int i) { throw new NotSupportedException(); }
        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length) { throw new NotSupportedException(); }
        public char GetChar(int i) { throw new NotSupportedException(); }
        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length) { throw new NotSupportedException(); }
        public Guid GetGuid(int i) { throw new NotSupportedException(); }
        public Int16 GetInt16(int i) { throw new NotSupportedException(); }
        public Int32 GetInt32(int i) { throw new NotSupportedException(); }
        public Int64 GetInt64(int i) { throw new NotSupportedException(); }
        public float GetFloat(int i) { throw new NotSupportedException(); }
        public double GetDouble(int i) { throw new NotSupportedException(); }
        public String GetString(int i) { throw new NotSupportedException(); }
        public Decimal GetDecimal(int i) { throw new NotSupportedException(); }
        public DateTime GetDateTime(int i) { throw new NotSupportedException(); }
        public IDataReader GetData(int i) { throw new NotSupportedException("GetData not supported."); }
        public bool IsDBNull(int i) { throw new NotSupportedException(); }
        void IDisposable.Dispose()
        {
            System.GC.SuppressFinalize(this);
        }

    }

}
