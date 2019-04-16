using ExcelToSqlBulkCopy.Properties;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Input;
using System.Windows.Threading;

namespace ExcelToSqlBulkCopy
{
    class MainWindowsViewModel : INotifyPropertyChanged
    {
        private string _excelFileName;
        private string _excelSheetName;
        private int _progressbarMaximum;
        private int _probressbarValue;
        private string _log;
        private bool _isEnabled;
        private string _destinationTableName;
        private string[] _excelSheetList;
        private string[] _tableNames;
        private Dispatcher _uiThread;
        private string _backuptable;

        public ICommand OpenExcelFileCommand => new RelayCommand(OpenExcelFile);
        public ICommand StartCommand => new RelayCommand(Start);

        public MainWindowsViewModel()
        {
            _excelFileName = Settings.Default.ExcelFileName;
            _excelSheetName = Settings.Default.ExcelSheet;
            
            var list = new List<string>();
            if (Settings.Default.ExcelSheetList != null)
            {
                foreach (var item in Settings.Default.ExcelSheetList)
                {
                    list.Add(item);
                }
            }
            _excelSheetList = list.ToArray();

            _isEnabled = true;
            _uiThread = Dispatcher.CurrentDispatcher;
            var del = new Action(LoadTableNameList);
            del.BeginInvoke(null, null);
        }

        private void LoadTableNameList()
        {
            IsEnabled = false;
            Log = "Loading table name list. ";
            try
            {
                using (SqlConnection destinationConnection = new SqlConnection(Settings.Default.SqlConnectionString))
                {
                    destinationConnection.Open();
                    var tableNames = GetTableNames(destinationConnection);
                    _uiThread.Invoke(new Action(() =>
                    {
                        TableNames = tableNames;
                        DestinationTableName = Settings.Default.DestinationTableName;
                    }
                    ));
                }

                Log += "Complete.";
                Log += Environment.NewLine + Environment.NewLine + "Before loading any data from Excel, " 
                    + Environment.NewLine + "Please ensure that the original column names are in the top row"
                    + Environment.NewLine + "and all blank rows have been removed."
                    + Environment.NewLine +"Customer/Supplier data : please remove any rows that do not have a Transaction Number. " 
                    + Environment.NewLine + "NRS data : please ensure that there are no duplicate Transaction numbers. ";
            }
            catch (Exception ex)
            {
                Log = "Error loading table names: " + ex.ToString();
            }
            finally
            {
                IsEnabled = true;
            }
        }

        public string[] TableNames
        {
            get => _tableNames;
            set
            {
                _tableNames = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TableNames)));
            }
        }

        public string[] ExcelSheetList
        {
            get => _excelSheetList;
            set
            {
                _excelSheetList = value;
                Settings.Default.ExcelSheetList = new StringCollection();
                Settings.Default.ExcelSheetList.AddRange(value);
                Settings.Default.Save();
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ExcelSheetList)));
            }
        }

        public string DestinationTableName
        {
            get => _destinationTableName;
            set
            {
                _destinationTableName = value;
                Settings.Default.DestinationTableName = _destinationTableName;
                Settings.Default.Save();
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DestinationTableName)));
            }
        }

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                _isEnabled = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsEnabled)));
            }
        }

        public int ProgressbarMaximum
        {
            get => _progressbarMaximum;
            set
            {
                if (_progressbarMaximum != value)
                {
                    _progressbarMaximum = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ProgressbarMaximum)));
                }
            }
        }

        public int ProgressbarValue
        {
            get => _probressbarValue;
            set
            {
                if (_probressbarValue != value)
                {
                    _probressbarValue = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ProgressbarValue)));
                }
            }
        }

        public string Log
        {
            get => _log; set
            {
                _log = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Log)));
            }
        }

        public string ExcelFileName
        {
            get => _excelFileName;
            set
            {
                _excelFileName = value;
                Settings.Default.ExcelFileName = value;
                Settings.Default.Save();
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ExcelFileName)));
            }
        }

        public string ExcelSheet
        {
            get => _excelSheetName;
            set
            {
                _excelSheetName = value;
                Settings.Default.ExcelSheet = value;
                Settings.Default.Save();
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ExcelSheet)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void Start(object param)
        {
            IsEnabled = false;
            var action = new Action(() => RunBulkCopy(_excelFileName, _excelSheetName, _destinationTableName, true));
            action.BeginInvoke(null, null);
        }

        private void OpenExcelFile(object param)
        {
            var dialog = new OpenFileDialog();
            string fileName = Settings.Default.ExcelFileName;

            if (!string.IsNullOrEmpty(fileName))
            {
                dialog.InitialDirectory = Path.GetDirectoryName(fileName);
                dialog.FileName = fileName;
            }

            dialog.Filter = "Excel files (*.xlsx) | *.xlsx";
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == true)
            {
                ExcelFileName = dialog.FileName;
                ExcelSheetList = GetWorksheetNames(ExcelFileName);
                ExcelSheet = ExcelSheetList.FirstOrDefault();
            }
        }






        private string[] GetTableNames(SqlConnection connection)
        {

            //AW Added hard code list to test
            List<string> GetTableNamesList = new List<string>();

            //Add items
            GetTableNamesList.Add("NominalUpload");
            GetTableNamesList.Add("VSL_NominalUpload");
            GetTableNamesList.Add("SupplierUpload_SSG");
            GetTableNamesList.Add("CustomerUpload_SSG");
            GetTableNamesList.Add("NRS_Upload");
            GetTableNamesList.Add("BudgetUpload");
            GetTableNamesList.Add("NRS_UploadTest");
            GetTableNamesList.Add("NominalUploadTest");
            
            //GetTableNamesList.Add("NominalUploadTest");

            var myArray = GetTableNamesList.ToArray();
            return myArray;

            //using (SqlCommand cmd = connection.CreateCommand())
            //{
            //    cmd.CommandText = "SELECT TABLE_SCHEMA + '.' + TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_SCHEMA, TABLE_NAME";
            //    cmd.CommandType = CommandType.Text;
            //    var reader = cmd.ExecuteReader();
            //    return GetColumnmNamesFromDataReader(reader);
            //}
        }

        private string[] GetSqlColumnNames(SqlConnection connection, string tableName)
        {
            using (SqlCommand cmd = connection.CreateCommand())
            {
                // the ; replacement is a rudimentary way to avoid a sql injection attack
                cmd.CommandText = "SELECT TOP 1 * FROM " + tableName.Replace(";", "");
                cmd.CommandType = CommandType.Text;
                var reader = cmd.ExecuteReader();
                return GetColumnmNamesFromDataReader(reader);
            }
        }

        //AW Added to backup table before appending
 
        private string CallBackupTable (SqlConnection connection, string tableName)
        {
            using (SqlCommand cmd = connection.CreateCommand())
            {
                // the ; replacement is a rudimentary way to avoid a sql injection attack
                //cmd.CommandText = "[dbo].[uspCopyNominal]";
                cmd.CommandText = "[dbo].[uspCopyNominal]";
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.CreateParameter("@FromTableName", adVarChar, adParamInput, 40, strFromTableName);
                SqlParameter param;
                param = cmd.Parameters.Add("@FromTableName", SqlDbType.NVarChar, 40);
                param.Value = tableName;
                param = cmd.Parameters.Add("@NewTableName", SqlDbType.NVarChar, 40);
                param.Value = tableName + DateTime.Now.ToString("dd_mm_yy_hhmmss");


                //cmd.CreateParameter("@FromTableName", adVarChar, adParamInput, 40, strFromTableName);
                //cmd.Parameters.Add("@NewTableName", SqlDbType.NVarChar, 40, "NRS_Upload_bktest");
                //cmd.Parameters.Add("@FromTableName", SqlDbType.NVarChar, 40, tableName);
                connection.Open();
                var reader = cmd.ExecuteReader();
                connection.Close();
                return "Backup Complete";
                
            }
        }

        private string[] GetColumnmNamesFromDataReader(IDataReader dataReader)
        {
            string[] columns = new string[dataReader.FieldCount];
            for (int i = 0; i < columns.Length; i++)
            {
                columns[i] = dataReader.GetName(i);
            }

            return columns;
        }

        private string[] GetWorksheetNames(string fileName)
        {
            Log = "Opening excel file";
            try
            {
                using (var excelPackage = new ExcelPackage())
                using (var stream = File.OpenRead(fileName))
                {
                    excelPackage.Load(stream);
                    string[] list = excelPackage.Workbook.Worksheets.Select(x => x.Name).ToArray();
                    Log = "Opening excel file. Success.";
                    return list;
                }
            }
            catch (Exception ex)
            {
                Log = "Error: " + ex.Message;
                return new string[0];
            }
        }

        private void RunBulkCopy(string fileName, string sheetName, string destinationTableName, bool hasHeader = true)
        {
            try
            {
                // run some checks
                if (string.IsNullOrEmpty(fileName))
                {
                    Log = "Excel file name cannot be empty";
                    return;
                }
                else if (string.IsNullOrEmpty(sheetName))
                {
                    Log = "Excel sheet name cannot be empty";
                    return;
                }
                else if (string.IsNullOrEmpty(destinationTableName))
                {
                    Log = "Destination table name cannot be empty";
                    return;
                }

                try
                {
                    Log = "Reading excel file: " + fileName;
                    using (var sourceReader = new ExcelDataReader(fileName, sheetName, hasHeader))
                    {
                        sourceReader.ReadProgress += SourceReader_ReadProgress;
                        string[] excelColumnNames = GetColumnmNamesFromDataReader(sourceReader);

                        // gets the sql column names from the destination table
                        Log = "Connecting to sql server";
                        string[] sqlColumnNames;
                        using (SqlConnection destinationConnection = new SqlConnection(Settings.Default.SqlConnectionString))
                        {
                            destinationConnection.Open();
                            string tableName = destinationTableName;
                            sqlColumnNames = GetSqlColumnNames(destinationConnection, tableName);
                        }

                        // this needs to be a new sql connection for some reason
                        using (SqlConnection destinationConnection = new SqlConnection(Settings.Default.SqlConnectionString))
                        {


                            //AW Test calling procedure to backup table
                            _backuptable = CallBackupTable(destinationConnection, DestinationTableName);

                            destinationConnection.Open();
                            Log = "Bulk copying data";
                            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(destinationConnection))
                            {
                                bulkCopy.DestinationTableName = destinationTableName;
                                bulkCopy.BatchSize = 500; // number of rows to copy in a batch
                                bulkCopy.BulkCopyTimeout = (int)(TimeSpan.FromMinutes(60).TotalSeconds); // 1 hour timeout

                                // add mapping where the column names match exactly
                                foreach (string columnName in excelColumnNames.Intersect(sqlColumnNames))
                                {
                                    bulkCopy.ColumnMappings.Add(new SqlBulkCopyColumnMapping(columnName, columnName));
                                }

                                bulkCopy.WriteToServer(sourceReader);
                            }
                        }
                    }
                }
                catch (IOException ex)
                {
                    Log = "Error: " + ex.Message;
                }
                catch (Exception ex)
                {
                    Log = "Error: " + ex.ToString();
                }
            }
            finally
            {
                IsEnabled = true;
            }
        }

        private void SourceReader_ReadProgress(object sender, ExcelDataReaderProgressEventArgs e)
        {
            ProgressbarValue = e.NumRecordsRead;
            ProgressbarMaximum = e.TotalNumRecords;
            Log = $"Copied {e.NumRecordsRead:#,##0} of {e.TotalNumRecords:#,##0} records. ";

            if (e.NumRecordsRead == e.TotalNumRecords)
            {
                Log += "Complete.";
            }
        }
    }
}
