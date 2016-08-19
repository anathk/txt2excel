using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace txt2excel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }
        public string txtFilePath { set; get; }
        public string excelFilePath { set; get; }

        public int intervalNum = 0;
        private int textFileLineCount = 0;
        private int counter = 0;



        private void LoadTxtFileBtn_OnClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                txtPathTextBox.Text = openFileDialog.FileName;
                txtFilePath = openFileDialog.FileName;
                excelFilePath = txtFilePath.Replace("txt", "xls");
                excelPathTextBox.Text = excelFilePath;
                textFileLineCount = File.ReadLines(txtFilePath).Count();
            }
                //txtEditor.Text = File.ReadAllText(openFileDialog.FileName);
        }

        private void LoadExcelBtn_OnClick(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void StartBtn_OnClick(object sender, RoutedEventArgs e)
        {
            startBtn.IsEnabled = false;
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += writeToExcelTask;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerAsync();
            //Task.Factory.StartNew(writeToExcelTask);

        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //progressBar.Value = e.ProgressPercentage;
            progressBar.Value = (((Double)counter)/((Double)textFileLineCount))*100;
            //progressBarStatusText.Text = progressBar.Value.ToString();
            progressBarStatusText.Text = string.Format("{0:P2}", progressBar.Value/100);
        }

        private void writeToExcelTask(object sender, DoWorkEventArgs e)
        {
            StreamReader txtFile = new StreamReader(txtFilePath);
            string line = "";           
            int writeCounter = 1;

            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            var oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            var oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            object misValue = System.Reflection.Missing.Value;

            for (int i = 0; i < 16; i++)
            {
                txtFile.ReadLine();
                counter++;
            }

            while ((line = txtFile.ReadLine()) != null)
            {

                List<string> tempRowData = new List<string>();
                tempRowData.AddRange(line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries));

                for (int col = 0; col < tempRowData.Count; col++)
                {
                    oSheet.Cells[writeCounter + 16, col + 1] = tempRowData[col];

                }
                writeCounter++;


                if (intervalNum > 0)
                {
                    for (int i = 0; i < intervalNum; i++)
                    {
                        txtFile.ReadLine();
                        counter++;
                    }
                }

                
                counter++;
                (sender as BackgroundWorker).ReportProgress(counter / textFileLineCount * 100);
                //progressBar.Value = counter/textFileLineCount*100;
            }

            oWB.SaveAs(excelFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            oWB.Close(true, misValue, misValue);
            oXL.Quit();

            releaseObject(oSheet);
            releaseObject(oWB);
            releaseObject(oXL);
            startBtn.IsEnabled = true;
            MessageBox.Show("Dang Dang Dang!!!");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void IntervalNumberTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            String intervalStr = intervalNumberTextBox.Text;
            try
            {
                intervalNum = Int32.Parse(intervalStr);
            }
            catch (FormatException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
