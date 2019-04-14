using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToSqlBulkCopy
{
    public class ExcelDataReaderProgressEventArgs : EventArgs
    {
        public int NumRecordsRead { get; private set; }
        public int TotalNumRecords { get; private set; }

        public ExcelDataReaderProgressEventArgs(int numRecordsRead, int totalNumRecords)
        {
            NumRecordsRead = numRecordsRead;
            TotalNumRecords = totalNumRecords;
        }
    }
}
