using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAppWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Excel.Application _excel = null;
        private Excel.Workbook _wb = null;
        private string _workBookFileExtension = null;
        public MainWindow()
        {
            _excel = Application.Current.Resources["Excel"] as Excel.Application;
            InitializeComponent();
        }
        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            try { _wb?.Close(false); }
            catch { }
            TargetFile.Clear();
            Result.Clear();
            BtnRunProgram.IsEnabled = false;
            BtnSaveFile.IsEnabled = false;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel-compatible files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm"; 
            if (openFileDialog.ShowDialog() == true)
            {
                TargetFile.Text = openFileDialog.FileName;
                string[] filenameSplits = openFileDialog.FileName.Split('.');
                _workBookFileExtension =filenameSplits[filenameSplits.Length - 1];
                BtnRunProgram.IsEnabled = true;
            }
        }
        private void TextBox_Drop(object sender, DragEventArgs e)
        {
            try { _wb?.Close(false); }
            catch { }
            TargetFile.Clear();
            Result.Clear();
            BtnRunProgram.IsEnabled = false;
            BtnSaveFile.IsEnabled = false;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0)
                {
                    TargetFile.Text = files[0];
                    string[] filenameSplits = files[0].Split('.');
                    _workBookFileExtension = filenameSplits[filenameSplits.Length - 1];
                    BtnRunProgram.IsEnabled = true;
                }
            }
        }
        private void TextBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }
        private void runProgram(object sender, EventArgs e)
        {
            Result.Clear();
            Result.AppendText("Working...");
            try
            {
                _wb = _excel?.Workbooks.Open(TargetFile.Text, false, true);
                Array olinks = _wb.LinkSources(Excel.XlLink.xlExcelLinks) as Array;
                if (olinks != null)
                {
                    for (int i = 1; i <= olinks.Length; i++)
                    {
                        Result.AppendText(Environment.NewLine+"Link detected: "+ (string)olinks.GetValue(i));
                        _wb.BreakLink((string)olinks.GetValue(i), Excel.XlLinkType.xlLinkTypeExcelLinks);
                    }
                }
                Result.AppendText(Environment.NewLine + "Finished Success!");
                BtnSaveFile.IsEnabled = true;
            }
            catch (Exception ex)
            {
                Result.AppendText(Environment.NewLine + "ERROR: " + ex.Message);
            }
            finally
            {

            }
        }
        private void saveFile(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = $"Excel file (*.{_workBookFileExtension})|*.{_workBookFileExtension}";
            if (saveFileDialog.ShowDialog() == true)
            {
                _wb?.SaveAs(saveFileDialog.FileName);
                Result.AppendText(Environment.NewLine + "File saved as: " + saveFileDialog.FileName);
                _wb?.Close(false);
                BtnSaveFile.IsEnabled = false;
            }
        }

    }
}
