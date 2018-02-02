// <copyright file="ImportDataReader.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
namespace SqlToCsv
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;

    /// <summary>
    /// Import data reader class.
    /// </summary>
    internal class ImportDataReader : IDataReader
    {
        private string[] fields;
        private char fieldSeparator;
        private TextReader reader;
        private KeyValuePair<string, string>[] replace;
        private bool trimWhiteSpace;
        private string[] values;

        public ImportDataReader(TextReader reader, char fieldSeparator, bool trimWhiteSpace, KeyValuePair<string, string>[] replace)
        {
            this.reader = reader;
            this.fieldSeparator = fieldSeparator;
            this.trimWhiteSpace = trimWhiteSpace;
            this.replace = replace;

            this.fields = CSV.Import(reader, this.fieldSeparator, this.trimWhiteSpace, replace);
        }

        public int Depth
        {
            get { throw new Exception("The method or operation is not implemented."); }
        }

        public int FieldCount
        {
            get { return this.fields.Length; }
        }

        public bool IsClosed
        {
            get { throw new Exception("The method or operation is not implemented."); }
        }

        public int RecordsAffected
        {
            get { throw new Exception("The method or operation is not implemented."); }
        }

        public object this[string name]
        {
            get { throw new Exception("The method or operation is not implemented."); }
        }

        public object this[int i]
        {
            get { throw new Exception("The method or operation is not implemented."); }
        }

        public void Close()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void Dispose()
        {
        }

        public bool GetBoolean(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public byte GetByte(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public char GetChar(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public IDataReader GetData(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public string GetDataTypeName(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public DateTime GetDateTime(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public decimal GetDecimal(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public double GetDouble(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public Type GetFieldType(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public float GetFloat(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public Guid GetGuid(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public short GetInt16(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public int GetInt32(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public long GetInt64(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public string GetName(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public int GetOrdinal(string name)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public DataTable GetSchemaTable()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public string GetString(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public object GetValue(int i)
        {
            return this.values[i];
        }

        public int GetValues(object[] values)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public bool IsDBNull(int i)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public bool NextResult()
        {
            return false;
        }

        public bool Read()
        {
            this.values = CSV.Import(this.reader, this.fieldSeparator, this.trimWhiteSpace, this.replace);
            return this.values != null;
        }
    }
}