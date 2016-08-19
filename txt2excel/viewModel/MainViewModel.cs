using System.IO;
using System.Linq;
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

        private int counter = 0;


        public RelayCommand LoadFileCommand { get; set; }
        public RelayCommand StartCommand { get; set; }


        /// <summary>
        /// Initializes a new instance of the MainViewModel class.
        /// </summary>
        public MainViewModel()
        {
            LoadFileCommand = new RelayCommand(LoadFile, CanLoadFileExecute);
            
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
    }
}