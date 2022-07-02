using ExcelDataReader;
using MyWpfCoreExcelImport.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Linq;
using System.Windows.Controls;

namespace MyWpfCoreExcelImport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// we need:
    /// Install-Package ExcelDataReader
    /// Install-Package ExcelDataReader.DataSet
    /// Install-Package System.Data.SqlClient
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public delegate void VisibilityChangeEvent(object sender, ChangeEventArgs e);
        public event PropertyChangedEventHandler PropertyChanged;

        private VisibilityChangeEvent _visibilityChangeEvent;
        private Tools.ToolWindow _toolWindow;
        private DataSet _ds;
        private string _connectionString;

        #region INotify Changed Properties  
        private string message;
        public string Message
        {
            get { return message; }
            set { SetField(ref message, value, nameof(Message)); }
        }

        private ObservableCollection<string> itemList;
        public ObservableCollection<string> ItemList
        {
            get { return itemList; }
            set { SetField(ref this.itemList, value, nameof(ItemList)); }
        }
        private string newItem;
        public string NewItem
        {
            get
            {
                return newItem;
            }
            set
            {
                if (newItem != value)
                {
                    newItem = value;
                    var item = ItemList.SingleOrDefault(x => x == newItem);
                    if (item == null)
                        ItemList.Insert(0, newItem);
                    SelectedItem = newItem;
                }
            }
        }
        private string selectedItem;
        public string SelectedItem
        {
            get
            {
                return selectedItem;
            }
            set
            {
                if (selectedItem != value)
                {
                    selectedItem = value;
                    if (selectedItem == ItemListToXml.DELETE_COMMAND)
                    {
                        ItemList.Clear();
                        ItemList.Add(ItemListToXml.DELETE_COMMAND);
                    }
                    SetField(ref this.selectedItem, value, nameof(SelectedItem));
                }
            }
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

        private ObservableCollection<string> tableList;
        public ObservableCollection<string> TableList
        {
            get { return tableList; }
            set { SetField(ref this.tableList, value, nameof(TableList)); }
        }
        private string selectedComboBoxTableName;
        public string SelectedComboBoxTableName
        {
            get { return selectedComboBoxTableName; }
            set { SetField(ref selectedComboBoxTableName, value, nameof(SelectedComboBoxTableName)); }
        }
        private int selectedTableIndex;
        public int SelectedTableIndex
        {
            get { return selectedTableIndex; }
            set { SetField(ref selectedTableIndex, value, nameof(SelectedTableIndex)); }
        }
        private string selectedTable;
        public string SelectedTable
        {
            get
            {
                return selectedTable;
            }
            set
            {
                if (selectedTable != value)
                {
                    selectedTable = value;
                    if (selectedTable == ItemListToXml.DELETE_COMMAND)
                    {
                        TableList.Clear();
                        TableList.Add(ItemListToXml.DELETE_COMMAND);
                    }
                    SetField(ref this.selectedTable, value, nameof(SelectedTable));
                }
            }
        }
        #endregion

        SqlConnection Connection { get; set; }
        public ItemListToXml ItemListToXml { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            DataContext = this;
            Message = "";

#if DEBUG
            Title += "    Debug Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
#else
            Title += "    Release Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
#endif
            // Necessary for ExcelDataReader and .Net Core
            Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            _visibilityChangeEvent = new VisibilityChangeEvent(VisibilityChangeEventHandler);

            ItemListToXml = new ItemListToXml();
            ItemList = ItemListToXml.Load(ref selectedItem);

            TableList = new ObservableCollection<string>();
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
        /// Button_1_Click
        /// For testing and debugging
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_1_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Button_1_Click");
            Console.Beep();
            Message = "You pressed Button #1 Beep :)";
        }

        /// <summary>
        /// Button_ChangeLog_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_ChangeLog_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Button_1_Click");

            ChangeLogCoreUtilityDll.ChangeLogTxtToolWindow ChangeLogTxtToolWindow = new ChangeLogCoreUtilityDll.ChangeLogTxtToolWindow(this);
            ChangeLogTxtToolWindow.ShowChangeLogWindow("ChangeLog.txt");
            Message = "ChangeLog";
        }

        /// <summary>
        /// Button_AddDefaultConnection_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_AddDefaultConnection_Click(object sender, RoutedEventArgs e)
        {
            var p = Properties.Settings.Default;
            NewItem = p.DefaultConnectionString;
            Message = "Add default Connection Item to List";
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
            SelectedTableIndex = 0;
            ClearDataSet();

            try
            {
                DataSet ds = ReadExcelFile(excelFile);
                _ds = ds;
                if (_ds?.Tables.Count > 0)
                {
                    SelectedComboBoxTableName = _ds.Tables[SelectedTableIndex].TableName;
                    myDataGrid.DataContext = _ds.Tables[SelectedTableIndex];
                    String createTableStatment = CreateTABLE(_ds.Tables[SelectedTableIndex]);
                    _toolWindow!.CreateTableStatment = createTableStatment;
                    Debug.WriteLine(createTableStatment);
                }
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
                {
                    if(!UploadDataSet(_ds))
                    {
                        Console.Beep();
                        Console.Beep();
                        Message = $"Database update fail";
                        MessageBox.Show($"Data has not yet been imported !!", Title, MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
                else
                {
                    Console.Beep();
                    Console.Beep();
                    Message = $"Database update fail";
                    MessageBox.Show($"Data has not yet been imported, ds == null !!", Title, MessageBoxButton.OK, MessageBoxImage.Error);
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

            Message = $"Database updated {_ds.Tables[SelectedTableIndex].Rows.Count} Rows added";
            Console.Beep();

            String createTableStatment = CreateTABLE(_ds.Tables[SelectedTableIndex]);
            Debug.WriteLine(createTableStatment);
        }

        /// <summary>
        /// Button_Clear_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Clear_Click(object sender, RoutedEventArgs e)
        {
            ClearDataSet();
        }
        private void ClearDataSet()
        {
            _ds?.Clear();
            _ds = null;
            TableList?.Clear();
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
            {
                _toolWindow = CreateToolWindow();
                _toolWindow.VisibilityChange += _visibilityChangeEvent;
            }
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
        /// ComboBox_KeyDown
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Return)
            {
                string newItemValue = ((System.Windows.Controls.TextBox)e.OriginalSource).Text;
                var item = ItemList.SingleOrDefault(x => x == newItemValue);
                if (item == null)
                    ItemList.Insert(0, newItemValue);
            }
        }

        /// <summary>
        /// ComboBox_SelectionChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            string text = "";
            int index;

            try { text = e.AddedItems[0] as string; } catch { }
            index = (sender as ComboBox).SelectedIndex;

            if (_ds != null && _ds.Tables.Count > 0)
            {
                myDataGrid.DataContext = _ds.Tables[index];
                _toolWindow!.CreateTableStatment = CreateTABLE(_ds.Tables[SelectedTableIndex]);
            }

            Debug.WriteLine($"ComboBox_SelectionChanged text={text} index={index}");
        }

        /// <summary>
        /// VisibilityChangeEventHandler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VisibilityChangeEventHandler(object sender, ChangeEventArgs e)
        {
            ShowSQLHelpWindow = e.IsVisible;
            Debug.WriteLine($"VisibilityChangeEventHandler {e.IsVisible}");
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
            ItemListToXml.Save(SelectedItem, ItemList);
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
                _connectionString = SelectedItem;
                p.DBName = GetDBName(_connectionString);
                Connection = new SqlConnection(_connectionString);
                Connection.Open();
                Console.Beep();
                Message = $"Connection to: {Connection!.ConnectionString}  - open";
            }
            catch (Exception ex)
            {
                Console.Beep();
                Console.Beep();
                if(Connection != null)
                    Message = $"Connection to: {Connection!.ConnectionString}  - failed";
                else
                    Message = $"Connection to DB  - failed";
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

            dataSetConfiguration = new ExcelDataSetConfiguration()
            {
                // Gets or sets a value indicating whether to set the DataColumn.DataType 
                // property in a second pass.
                UseColumnDataType = true,

                // Gets or sets a callback to determine whether to include the current sheet
                // in the DataSet. Called once per sheet before ConfigureDataTable.
                FilterSheet = (tableReader, sheetIndex) => true,

                // Gets or sets a callback to obtain configuration options for a DataTable. 
                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                {
                    // Gets or sets a value indicating the prefix of generated column names.
                    EmptyColumnNamePrefix = "Column",

                    // Gets or sets a value indicating whether to use a row from the 
                    // data as column names.
                    UseHeaderRow = true,
                }
            };


            p.DBName = GetDBName(_connectionString);
            using FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream, readerConfiguration);
            result = reader.AsDataSet(dataSetConfiguration);

            for(int i=0;i < result.Tables.Count;i++)
                TableList.Add(result.Tables[i].TableName);

            SelectedTableIndex = 0;
            if (result.Tables.Count > 0)
                SelectedTable = result.Tables[SelectedTableIndex].TableName;

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
                bulkCopy.DestinationTableName = ds.Tables[SelectedTableIndex].TableName;
                bulkCopy.WriteToServer(ds.Tables[SelectedTableIndex]);
                Message = $"Updating data in table {0} ds.Tables[SelectedTableIndex].TableName Success";

                return true;
            }
            catch (Exception ex)
            {
                Message = $"Failed to update table {0}. Error {1} ds.Tables[SelectedTableIndex].TableName, ex Failed";
                Debug.WriteLine(ex.ToString());

                return false;
            }
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
            sqlsc += $"DROP TABLE IF EXISTS {table.TableName};\n";
            sqlsc += $"CREATE TABLE {table.TableName} (";
            for (int i = 0; i < table.Columns.Count; i++)
            {
                sqlsc += "\n [" + table.Columns[i].ColumnName + "] ";
                string columnType = table.Columns[i].DataType.ToString();
                switch (columnType)
                {
                    case "System.Double":
                        if (table.Columns[i].ColumnName.ToLower() == "id")
                            sqlsc += " Int";
                        else
                            sqlsc += " Real";
                        break;
                    case "System.Decimal":
                        sqlsc += " Real";
                        break;
                    case "System.Int32":
                        sqlsc += " Int";
                        break;
                    case "System.Int64":
                        sqlsc += " Bigint";
                        break;
                    case "System.Int16":
                        sqlsc += " Smallint";
                        break;
                    case "System.Byte":
                        sqlsc += " Tinyint";
                        break;
                    case "System.DateTime":
                        sqlsc += " Datetime";
                        break;
                    case "System.String":
                    default:
                        sqlsc += string.Format(" Nvarchar({0})", table.Columns[i].MaxLength == -1 ? "max" : table.Columns[i].MaxLength.ToString());
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

    public class ChangeEventArgs : EventArgs
    {
        public bool IsVisible { get; set; }

        public ChangeEventArgs()
        {
        }

        public ChangeEventArgs(bool isVisible)
        {
            IsVisible = isVisible;
        }
    }
}
