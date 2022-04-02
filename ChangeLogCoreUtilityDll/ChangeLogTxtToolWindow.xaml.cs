using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace ChangeLogCoreUtilityDll
{
    /// <summary>
    /// Interaction logic for ChangeLogTxtToolWindow.xaml
    /// </summary>
    public partial class ChangeLogTxtToolWindow : Window
    {
        public string BackgroundDisplayText { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public ChangeLogTxtToolWindow(object parent)
        {
            InitializeComponent();
            ((ContentControl)parent).Unloaded += new RoutedEventHandler(HaldelParentClosing);
            BackgroundDisplayText = "CangeLog";
        }

        /******************************/
        /*       Button Events        */
        /******************************/
        #region Button Events

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
        /// ParentConsing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void HaldelParentClosing(object sender, EventArgs e)
        {
            Close();
        }

        #endregion
        /******************************/
        /*      Other Functions       */
        /******************************/
        #region Other Functions

        /// <summary>
        /// ShowChangeLogWindow
        /// </summary>
        /// <param name="filename"></param>
        public void ShowChangeLogWindow(string filename)
        {
            LableBackgroundDisplayText.Content = BackgroundDisplayText;
            Paragraph Paragraph = new Paragraph();
            try
            {
                Paragraph.Inlines.Add(System.IO.File.ReadAllText(filename));
            }
            catch (Exception)
            {
                Paragraph.Inlines.Add("File Not Found");
            }
            Paragraph.FontFamily = new FontFamily("Courier New");
            Paragraph.FontSize = 12;
            FlowDocument Document = new FlowDocument(Paragraph);
            FlowDocReader.Document = Document;
            FlowDocReader.ViewingMode = FlowDocumentReaderViewingMode.Scroll;
            Show();
        }

        #endregion
    }
}
