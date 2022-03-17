using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace MyWpfCoreExcelImport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private Tools.ToolWindow _toolWindow;

        #region INotify Changed Properties  
        private string message;
        public string Message
        {
            get { return message; }
            set { SetField(ref message, value, nameof(Message)); }
        }
        private bool showSQLHelpWindow;
        public bool ShowSQLHelpWindow
        {
            get { return showSQLHelpWindow; }
            set
            {
                if (value)
                    ShowToolWindow();
                else
                    HideToolWindow();
                SetField(ref showSQLHelpWindow, value, nameof(ShowSQLHelpWindow));
            }
        }
        #endregion

        SqlConnection Connection { get; set; }

        DataSet _ds;

        /// <summary>
        /// Constructor
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            DataContext = this;

#if DEBUG
            Title += "    Debug Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
#else
            Title += "    Release Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
#endif
            // Necessary for ExcelDataReader and .Net Core
            Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        /******************************/
        /*       Button Events        */
        /******************************/
        #region Button Events

        /// <summary>
        /// Button_Connect_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Connect_Click(object sender, RoutedEventArgs e)
        {
            ConnectToDB();
        }

        /// <summary>
        /// Button_SetExcelFile_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_SetExcelFile_Click(object sender, RoutedEventArgs e)
        {
            var p = Properties.Settings.Default;
            p.ExcelFile = GetFileFromDialog(p.ExcelFile, "Excel files|*.xlsx|Excel files (*.xls)|*.xls|All files (*.*)|*.*");
        }

        /// <summary>
        /// Button_Import_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Import_Click(object sender, RoutedEventArgs e)
        {
            String excelFile;
            var p = Properties.Settings.Default;

            excelFile = p.ExcelFile;

            try
            {
                DataSet ds = ReadExcelFile(excelFile);
                _ds = PrepareDataSetForImport(ds);
                myDataGrid.DataContext = _ds.Tables[0];
            }
            catch (Exception ex)
            {
                Console.Beep();
                Console.Beep();
                Message = $"{excelFile} import fail";
                MessageBox.Show(ex.ToString(), Title, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            Message = $"{excelFile} imported";
            Console.Beep();
        }

        /// <summary>
        /// Button_UpdatedDB_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_UpdatedDB_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_ds != null)
                    UploadDataSet(_ds);
                else
                {
                    Console.Beep();
                    Console.Beep();
                    Message = $"Database updat fail";
                    MessageBox.Show($"Data has not yet been imported, ds == null", Title, MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.Beep();
                Console.Beep();
                Message = $"Database updat fail";
                MessageBox.Show(ex.ToString(), Title, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            Message = $"Database updated {_ds.Tables[0].Rows.Count} Rows added";
            Console.Beep();

            String createTableStatment = CreateTABLE(_ds.Tables[0]);
            Debug.WriteLine(createTableStatment);
        }

        /// <summary>
        /// Button_Clear_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Clear_Click(object sender, RoutedEventArgs e)
        {
            _ds!.Clear();
        }

        /// <summary>
        /// Button_Close_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion
        /******************************/
        /*      Menu Events          */
        /******************************/
        #region Menu Events

        #endregion
        /******************************/
        /*      Other Events          */
        /******************************/
        #region Other Events

        /// <summary>
        /// Window_Loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ConnectToDB();
            if (_toolWindow == null)
                _toolWindow = CreateToolWindow();
        }

        /// <summary>
        /// Lable_Message_MouseDown
        /// Clear Message
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Lable_Message_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Message = "";
        }

        /// <summary>
        /// Window_Closing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            Properties.Settings.Default.Save();
            if (_toolWindow != null)
                _toolWindow.Close();
        }

        #endregion
        /******************************/
        /*      Other Functions       */
        /******************************/
        #region Other Functions

        /// <summary>
        /// ConnectToDB
        /// </summary>
        private void ConnectToDB()
        {
            var p = Properties.Settings.Default;

            try
            {
                p.DBName = GetDBName(p.ConnectionString);
                Connection = new SqlConnection(p.ConnectionString);
                Connection.Open();
                Console.Beep();
                Message = $"Connection to: {Connection!.ConnectionString}  - open";
            }
            catch (Exception ex)
            {
                Console.Beep();
                Console.Beep();
                Message = $"Connection to: {Connection!.ConnectionString}  - failed";
                MessageBox.Show(ex.Message, Title, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// GetDBName
        /// </summary>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private string GetDBName(string connectionString)
        {
            string dBName = "";

            var sCStrring = new System.Data.SqlClient.SqlConnectionStringBuilder(connectionString);

            dBName = sCStrring.InitialCatalog;

            return dBName;
        }

        /// <summary>
        /// ReadExcelFile
        /// The data examples in this project come from:
        /// https://www.mockaroo.com/
        /// we need:
        /// Install-Package ExcelDataReader
        /// Install-Package ExcelDataReader.DataSet
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="readerConfiguration"></param>
        /// <param name="dataSetConfiguration"></param>
        /// <returns></returns>
        private DataSet ReadExcelFile(string filePath, ExcelReaderConfiguration readerConfiguration = null, ExcelDataSetConfiguration dataSetConfiguration = null)
        {
            var p = Properties.Settings.Default;
            DataSet result;

            p.DBName = GetDBName(p.ConnectionString);
            using FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream, readerConfiguration);
            result = reader.AsDataSet(dataSetConfiguration);

            return result;
        }

        /// <summary>
        /// UploadDataSet
        /// Install-Package System.Data.SqlClient
        /// </summary>
        /// <param name="ds"></param>
        /// <returns></returns>
        private bool UploadDataSet(DataSet ds)
        {
            try
            {
                SqlBulkCopy bulkCopy = new SqlBulkCopy(Connection, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.UseInternalTransaction, null);
                bulkCopy.DestinationTableName = ds.Tables[0].TableName;
                bulkCopy.WriteToServer(ds.Tables[0]);
                Message = $"Updating data in table {0} ds.Tables[0].TableName Success";

                return true;
            }
            catch (Exception ex)
            {
                Message = $"Failed to update table {0}. Error {1} ds.Tables[0].TableName, ex Failed";
                Debug.WriteLine(ex.ToString());

                return false;
            }
        }

        /// <summary>
        /// PrepareDataSetForImport
        /// This function operates on the data set in such a way that the 
        /// first row becomes the column names of the data tables and then 
        /// the 1 row is removed
        /// </summary>
        /// <returns></returns>
        private DataSet PrepareDataSetForImport(DataSet ds)
        {
            DataSet pds = ds;

            try
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    for (int j = 0; j < ds.Tables[i].Columns.Count; j++)
                    {
                        DataColumn c = ds.Tables[i].Columns[j];
                        var r = pds.Tables[i].Rows[0];
                        c.ColumnName = r.ItemArray[j].ToString();
                    }
                    ds.Tables[i].Rows.RemoveAt(0);

                    ds.Tables[0].TableName = Properties.Settings.Default.TableName;
                    String createTableStatment = CreateTABLE(ds.Tables[0]);
                    _toolWindow!.CreateTableStatment = createTableStatment;
                    Debug.WriteLine(createTableStatment);
                }
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                Debug.WriteLine(ex.ToString());
            }

            return pds;
        }

        /// <summary>
        /// CreateTABLE
        /// From:
        /// https://stackoverflow.com/questions/1348712/creating-a-sql-server-table-from-a-c-sharp-datatable
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public string CreateTABLE(DataTable table)
        {
            var p = Properties.Settings.Default;
            string sqlsc;
            sqlsc = $"--CREATE DATABASE {p.DBName};\n";
            sqlsc += $"--USE master; ALTER DATABASE {p.DBName} SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE {p.DBName};\n";
            sqlsc += $"USE {p.DBName};\n";
            sqlsc += $"DROP TABLE IF EXISTS {p.DBName}.{table.TableName};\n";
            sqlsc += $"CREATE TABLE {table.TableName} (";
            for (int i = 0; i < table.Columns.Count; i++)
            {
                sqlsc += "\n [" + table.Columns[i].ColumnName + "] ";
                string columnType = table.Columns[i].DataType.ToString();
                switch (columnType)
                {
                    case "System.Int32":
                        sqlsc += " int ";
                        break;
                    case "System.Int64":
                        sqlsc += " bigint ";
                        break;
                    case "System.Int16":
                        sqlsc += " smallint";
                        break;
                    case "System.Byte":
                        sqlsc += " tinyint";
                        break;
                    case "System.Decimal":
                        sqlsc += " decimal ";
                        break;
                    case "System.DateTime":
                        sqlsc += " datetime ";
                        break;
                    case "System.String":
                    default:
                        sqlsc += string.Format(" nvarchar({0}) ", table.Columns[i].MaxLength == -1 ? "max" : table.Columns[i].MaxLength.ToString());
                        break;
                }
                if (table.Columns[i].AutoIncrement)
                    sqlsc += " IDENTITY(" + table.Columns[i].AutoIncrementSeed.ToString() + "," + table.Columns[i].AutoIncrementStep.ToString() + ") ";
                if (!table.Columns[i].AllowDBNull)
                    sqlsc += " NOT NULL ";
                sqlsc += ",";                
            }
            sqlsc += "\n);\n";
            sqlsc += $"SELECT * FROM {table.TableName};\n";
            return sqlsc;
        }

        /// <summary>
        /// GetFileFromDialog
        /// </summary>
        /// <param name="defaultPath"></param>
        /// <param name="filter"></param>
        /// <returns></returns>
        private string GetFileFromDialog(string defaultPath, string filter)
        {
            string fileName = defaultPath;

            System.Windows.Forms.OpenFileDialog ofDialog = new System.Windows.Forms.OpenFileDialog();
            ofDialog.Title = "Open Excel File";
            ofDialog.Filter = filter;
            ofDialog.InitialDirectory = Path.GetDirectoryName(defaultPath);
            System.Windows.Forms.DialogResult res = ofDialog.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK)
                fileName = ofDialog.FileName;

            return fileName;
        }

        /// <summary>
        /// CreateToolWindow
        /// </summary>
        /// <returns></returns>
        private Tools.ToolWindow CreateToolWindow()
        {
            Tools.ToolWindow toolWindow = new Tools.ToolWindow();
            toolWindow.WindowStyle = WindowStyle.ToolWindow;
            toolWindow.ShowInTaskbar = false;
            toolWindow.Owner = this;

            return toolWindow;
        }

        /// <summary>
        /// ShowToolWindow
        /// </summary>
        private void ShowToolWindow()
        {
            _toolWindow.Show();
        }

        /// <summary>
        /// HideToolWindow
        /// </summary>
        private void HideToolWindow()
        {
            _toolWindow.Hide();
        }

        /// <summary>
        /// SetField
        /// for INotify Changed Properties
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="field"></param>
        /// <param name="value"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        protected bool SetField<T>(ref T field, T value, string propertyName)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
        private void OnPropertyChanged(string p)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
        }

        #endregion
    }
}
