using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Reflection;
using MappingTest.Excel_Classes;
using MappingTest.RIPL_Classes;
using MappingTest.Other_Classes;

namespace MappingTest.Excel_Classes
{
    public partial class ExcelVariables : PropertyChangedBase
    {
        public static ObservableCollection<ExcelVariables> GetColumns(string fileName, string sheet)
        {
            //DataTable dtColumns = new DataTable();
            //dtColumns.Columns.Add("Columns", typeof(string));
            ObservableCollection<ExcelVariables> lstXcelVar = new ObservableCollection<ExcelVariables>();
            int num = 1;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(fileName.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);



            foreach (Excel.Worksheet exSheet in excelBook.Sheets)
            {
                if (exSheet.Name == sheet)
                {
                    num = exSheet.Index;

                }

            }

            Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(num);
            Excel.Range excelRange = excelSheet.UsedRange;
            int colCnt = 0;

            List<string> list = new List<string>();
            for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
            {
                string strColumn = "";
                strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                list.Add(strColumn);
            }
            foreach (var item in list)
            {
                ExcelVariables item2 = new ExcelVariables();
                item2.XcelVar = item.ToString();
                lstXcelVar.Add(item2);
            }

            excelBook.Close(true, null, null);
            excelApp.Quit();

            return lstXcelVar;
        }


        public string XcelVar { get; set; }
        public override string ToString()
        {
            return this.XcelVar;
        }
      
    }
}
