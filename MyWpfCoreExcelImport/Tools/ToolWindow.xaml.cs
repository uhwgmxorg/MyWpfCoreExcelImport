using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace MyWpfCoreExcelImport.Tools
{
    /// <summary>
    /// Interaction logic for ToolWindow.xaml
    /// </summary>
    public partial class ToolWindow : Window, INotifyPropertyChanged
    {
        public event MainWindow.VisibilityChangeEvent VisibilityChange;
        public event PropertyChangedEventHandler PropertyChanged;

        #region INotify Changed Properties  
        private string createTableStatment;
        public string CreateTableStatment
        {
            get { return createTableStatment!; }
            set { SetField(ref createTableStatment!, value, nameof(CreateTableStatment)); }
        }
        #endregion

        /// <summary>
        /// constructor
        /// </summary>
        public ToolWindow()
        {
            InitializeComponent();
            DataContext = this;

            CreateTableStatment = "The CREATE TABLE statement will appear here after\nImport of the Excel-File";
        }

        /******************************/
        /*       Button Events        */
        /******************************/
        #region Button Events

        /// <summary>
        /// Button_CopyPrivateKey_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_CopyCreateTableStatment_Click(object sender, RoutedEventArgs e)
        {
            var p = Properties.Settings.Default;

            try
            {
                Clipboard.SetDataObject(CreateTableStatment);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Exception: {ex.Message}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Button_DeleteCreateTableStatment_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_DeleteCreateTableStatment_Click(object sender, RoutedEventArgs e)
        {
            CreateTableStatment = "";
        }

        /// <summary>
        /// Button_Click_Close
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_Close(object sender, RoutedEventArgs e)
        {
            Hide();
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
        /// Window_IsVisibleChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            VisibilityChange?.Invoke(this, new ChangeEventArgs(IsVisible));
        }

        /// <summary>
        /// Window_Closing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
        }

        #endregion
        /******************************/
        /*      Other Functions       */
        /******************************/
        #region Other Functions

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
