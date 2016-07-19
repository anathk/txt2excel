using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GalaSoft.MvvmLight;

namespace txt2excel
{
    class MainViewModel : ViewModelBase
    {

        public string txtFilePath { set; get; }     //Text file Path.
        public string excelFilePath { set; get; }   //Excel file Path.

        public int intervalNum { set; get; }        //Pick line in text file every {intervalNum} lines.
        public int textFileLineCount { set; get; }  //Total lines in text file.
        public int counter { set; get; }            //Current line count.
        public int startLineNumer { set; get; }     //Start line number

        public MainViewModel()
        {
            
        }


    }
}
