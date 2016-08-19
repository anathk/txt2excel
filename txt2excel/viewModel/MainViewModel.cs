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
        }

        public int intervalNum = 0;
        private int textFileLineCount = 0;
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
                textFilePath = openFileDialog.FileName;

                excelFilePath = textFilePath.Replace("txt", "xls");

                textFileLineCount = File.ReadLines(txtFilePath).Count();
            }
        }
    }
}