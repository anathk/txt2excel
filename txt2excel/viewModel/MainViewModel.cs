using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Microsoft.Win32;

namespace txt2excel.ViewModel
{
    /// <summary>
    /// This class contains properties that the main View can data bind to.
    /// <para>
    /// Use the <strong>mvvminpc</strong> snippet to add bindable properties to this ViewModel.
    /// </para>
    /// <para>
    /// You can also use Blend to data bind with the tool's support.
    /// </para>
    /// <para>
    /// See http://www.galasoft.ch/mvvm
    /// </para>
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        private string textFilePath;
        public string TextFilePath
        {
            set
            {
                textFilePath = value; 
                RaisePropertyChanged("TextFilePath");
            }
            get { return textFilePath; }
        }

        private string excelFilePath;
        public string ExcelFilePath
        {
            set
            {
                excelFilePath = value;
                RaisePropertyChanged("ExcelFilePath");
            }
            get
            {
                return excelFilePath;
            }
        }

        private int startLine = 0;
        public int StartLine
        {
            set
            {
                startLine = value;
                RaisePropertyChanged("StartLine");
            }
            get { return startLine;}
        }

        private int intervalNum = 0;
        public int IntervalNum
        {
            set
            {
                intervalNum = value;
                RaisePropertyChanged("IntervalNum");
            }
            get { return intervalNum;}
        }

        private int totalLines = 0;
        public int TotalLines
        {
            set
            {
                totalLines = value;
                RaisePropertyChanged("TotalLines");
            }
            get { return totalLines; }
        }

        private string progressBarText;
        public string ProgressBarText
        {
            get { return progressBarText; }
            set
            {
                progressBarText = value;
                RaisePropertyChanged("ProgressBarText");
            }
        }

        private double progressBarValue;
        public double ProgressBarValue
        {
            set
            {
                progressBarValue = value;
                RaisePropertyChanged("ProgressBarValue");
            }
            get
            {
                return progressBarValue;
            }
        }

        private string currentOption;
        public string CurrentOption
        {
            set
            {
                currentOption = value;
                RaisePropertyChanged("CurrentOption");
            }
            get
            {
                return currentOption;
            }
        }

        public List<string> ProcessOption { get; set; }

        private int counter = 0;


        public RelayCommand LoadFileCommand { get; set; }
        public RelayCommand StartCommand { get; set; }


        /// <summary>
        /// Initializes a new instance of the MainViewModel class.
        /// </summary>
        public MainViewModel()
        {
            LoadFileCommand = new RelayCommand(LoadFile, CanLoadFileExecute);
            StartCommand = new RelayCommand(StartProcessing, CanStartProcessingExcute);
            ProcessOption = new List<string>();
            ProcessOption.Add("Read All");
            ProcessOption.Add("Read Partial");


        }

        private bool CanStartProcessingExcute()
        {
            return true;
        }

        private void StartProcessing()
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += writeToExcelTask;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerAsync(); ;
        }

        private bool CanLoadFileExecute()
        {
            return true;
        }

        private void LoadFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                TextFilePath = openFileDialog.FileName;

                ExcelFilePath = textFilePath.Replace("txt", "xls");

                TotalLines = File.ReadLines(textFilePath).Count();
            }
        }

        private void writeToExcelTask(object sender, DoWorkEventArgs e)
        {
            StreamReader txtFile = new StreamReader(TextFilePath);
            string line = "";
            int writeCounter = 1;

            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            var oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            var oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            object misValue = System.Reflection.Missing.Value;

            for (int i = 0; i < StartLine - 1; i++)
            {
                txtFile.ReadLine();
                counter++;
            }

            #region "Text process option 1"
            while ((line = txtFile.ReadLine()) != null)
            {
                if (line.Equals(""))
                {
                    line = txtFile.ReadLine();
                    counter++;
                }
                List<string> tempRowData = new List<string>();
                if (CurrentOption.Equals("Read All"))
                {                   
                    tempRowData.AddRange(line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries));

                    for (int col = 0; col < tempRowData.Count; col++)
                    {
                        oSheet.Cells[writeCounter + StartLine, col + 1] = tempRowData[col];
                    }
                    writeCounter++;
                }
                else if (CurrentOption.Equals("Read Partial"))
                {
                    tempRowData.AddRange(line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries));
                    var nextLine = txtFile.ReadLine();
                    if (nextLine != null)
                    {
                        tempRowData.AddRange(nextLine.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries));
                        counter++;
                    }
                    for (int col = 0; col < tempRowData.Count; col++)
                    {
                        oSheet.Cells[writeCounter + StartLine, col + 1] = tempRowData[col];
                    }
                    writeCounter++;
                }


                if (intervalNum > 0)
                {
                    for (int i = 0; i < intervalNum; i++)
                    {
                        if (CurrentOption.Equals("Read All"))
                        {
                            txtFile.ReadLine();
                            counter++;
                        }
                        else if (CurrentOption.Equals("Read Partial"))
                        {
                            txtFile.ReadLine();
                            txtFile.ReadLine();
                            counter += 2;
                        }
                        
                    }
                }


                counter++;
                (sender as BackgroundWorker).ReportProgress(counter / TotalLines * 100);
                //progressBar.Value = counter/textFileLineCount*100;
            }
#endregion
            oWB.SaveAs(excelFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            oWB.Close(true, misValue, misValue);
            oXL.Quit();

            releaseObject(oSheet);
            releaseObject(oWB);
            releaseObject(oXL);

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

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //progressBar.Value = e.ProgressPercentage;
            ProgressBarValue = (((Double)counter) / ((Double)TotalLines)) * 100;
            //progressBarStatusText.Text = progressBar.Value.ToString();
            ProgressBarText = string.Format("{0:P2}", ProgressBarValue / 100);
        }
    }
}