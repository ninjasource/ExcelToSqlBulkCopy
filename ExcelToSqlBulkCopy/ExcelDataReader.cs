using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelToSqlBulkCopy
{
    class ExcelDataReader : IDataReader
    {
        string[] _columnNames;
        ExcelPackage _excelPackage;
        ExcelWorksheet _sheet;
        bool _hasHeader;
        int _rowIndex;
        object[] _currentRow;
        public event EventHandler<ExcelDataReaderProgressEventArgs> ReadProgress;
        int _numRows;

        public ExcelDataReader(string fileName, string sheetName, bool hasHeader = true)
        {
            _excelPackage = new ExcelPackage();
            var columnNames = new List<string>();
            _hasHeader = hasHeader;
            using (var stream = File.OpenRead(fileName))
            {
                _excelPackage.Load(stream);
            }

            var sheet = _excelPackage.Workbook.Worksheets[sheetName];
            foreach (var firstRowCell in sheet.Cells[1, 1, 1, sheet.Dimension.End.Column])
            {
                columnNames.Add(hasHeader ? firstRowCell.Text : string.Format("Column{0}", firstRowCell.Start.Column));
            }

            _columnNames = columnNames.ToArray();
            _sheet = sheet;
            _numRows = _sheet.Dimension.End.Row - (_hasHeader ? 1 : 0);
        }

        public void Close()
        {
            _excelPackage?.Dispose();
            _excelPackage = null;
            _sheet = null;
        }

        public void Dispose()
        {
            Close();
        }

        public string GetName(int i)
        {
            return _columnNames[i];
        }

        public int GetOrdinal(string name)
        {
            for (int i = 0; i < _columnNames.Length; i++)
            {
                if (_columnNames[i] == name)
                {
                    return i;
                }
            }

            throw new Exception($"Column name {name} not found");
        }

        public bool Read()
        {
            var excelRowNumber = (_hasHeader ? 2 : 1) + _rowIndex;
            if (excelRowNumber <= _sheet.Dimension.End.Row)
            {
                var row = _sheet.Cells[excelRowNumber, 1, excelRowNumber, _sheet.Dimension.End.Column];
                _currentRow = new object[_columnNames.Length];
                foreach (var cell in row)
                {
                    int colIndex = cell.Start.Column - 1;
                    if (colIndex < _currentRow.Length)
                    {
                        _currentRow[cell.Start.Column - 1] = cell.Value;
                    }
                }

                _rowIndex++;

                // raise a progress event every 100 records
                if (_rowIndex % 100 == 0)
                {
                    ReadProgress?.Invoke(this, new ExcelDataReaderProgressEventArgs(_rowIndex, _numRows));
                }
                return true;
            }

            ReadProgress?.Invoke(this, new ExcelDataReaderProgressEventArgs(_rowIndex, _numRows));
            return false;
        }

        public object this[int i] => _currentRow[i];

        public object this[string name] => _currentRow[GetOrdinal(name)];

        public int Depth => 0;

        public bool IsClosed => _excelPackage == null;

        public int RecordsAffected => -1;

        public int FieldCount => _columnNames.Length;

        public string GetString(int i)
        {
            return _currentRow[i]?.ToString();
        }

        public object GetValue(int i)
        {
            return _currentRow[i];
        }

        public bool GetBoolean(int i)
        {
            throw new NotImplementedException();
        }

        public byte GetByte(int i)
        {
            throw new NotImplementedException();
        }

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public char GetChar(int i)
        {
            throw new NotImplementedException();
        }

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public IDataReader GetData(int i)
        {
            throw new NotImplementedException();
        }

        public string GetDataTypeName(int i)
        {
            throw new NotImplementedException();
        }

        public DateTime GetDateTime(int i)
        {
            throw new NotImplementedException();
        }

        public decimal GetDecimal(int i)
        {
            throw new NotImplementedException();
        }

        public double GetDouble(int i)
        {
            throw new NotImplementedException();
        }

        public Type GetFieldType(int i)
        {
            throw new NotImplementedException();
        }

        public float GetFloat(int i)
        {
            throw new NotImplementedException();
        }

        public Guid GetGuid(int i)
        {
            throw new NotImplementedException();
        }

        public short GetInt16(int i)
        {
            throw new NotImplementedException();
        }

        public int GetInt32(int i)
        {
            throw new NotImplementedException();
        }

        public long GetInt64(int i)
        {
            throw new NotImplementedException();
        }

        public DataTable GetSchemaTable()
        {
            throw new NotImplementedException();
        }

        public int GetValues(object[] values)
        {
            throw new NotImplementedException();
        }

        public bool IsDBNull(int i)
        {
            throw new NotImplementedException();
        }

        public bool NextResult()
        {
            throw new NotImplementedException();
        }
    }
}
