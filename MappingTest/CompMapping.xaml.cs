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
using System.Windows.Controls.Primitives;

namespace MappingTest
{
    /// <summary>
    /// Interaction logic for CompMapping.xaml
    /// </summary>
    public partial class CompMapping : Window
    {
        private List<Components> riplcomp = new List<Components>();
        private List<ExcelComps> sourcecomp = new List<ExcelComps>();
        private List<TotalComp> totalcomp = new List<TotalComp>();
        MainWindow origWindow;

        public ObservableCollection<RIPLComps> lstComps { get; set; }
        public static List<ComponentMappingClass> lstCompMapping { get; set; }
       

        public CompMapping(MainWindow incomingWindow)
        {
            
            InitializeComponent();
            origWindow = incomingWindow;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            lstComps = new ObservableCollection<RIPLComps>();
            lstComps = GetRIPLComps();
            //lstCompMapping = new List<ComponentMappingClass>();

            //foreach (DataRow row in dt.Rows)
            //{

            //    riplcomp.Add(new Components { component = row[0].ToString() });
            //    totalcomp.Add(new TotalComp { ripl = row[0].ToString() });
            //}
            //RIPLComp.ItemsSource = riplcomp;

            //string selectedFile = ((MainWindow)Application.Current.MainWindow).filePath.Text;

            var q1 = MainWindow.VMapList.Where(comp => comp.RIPLVarID == 100);


            string componentColumn = q1.ToString();

            
            List<ExcelComps> lstXcelComps = new List<ExcelComps>();
            lstXcelComps = GetSourceComps(MainWindow.selectedFile, MainWindow.selectedSheet, "Component");


            ComponentMapping.ItemsSource = lstXcelComps;
        }
        public  List<ExcelComps> GetSourceComps(string fileName, string sheet, string column)
        {
            List<ExcelComps> lstXcelComps = new List<ExcelComps>();
            
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(fileName.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Sheets[sheet];
            Excel.Range excelRange = excelSheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int columnCount = excelRange.Columns.Count;
            Excel.Range result = null;
            string Address = null;
            Excel.Range columns = excelSheet.Rows[1] as Excel.Range;
            //foreach (Excel.Range c in columns.Cells)
            //{
            //    if(c.Value == column)
            //    {

            //    }
            //}
            List<string> columnValue = new List<string>();

            result = columns.Find(What: column, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlWhole, SearchOrder: Excel.XlSearchOrder.xlByColumns);
            Excel.Range cRng = null;
            if (result != null)
            {
                Address = result.Address;

                do
                {
                    for (int i = 2; i <= rowCount; i++)
                    {
                        cRng = excelSheet.Cells[i, result.Column] as Excel.Range;
                        if (cRng.Value != null)
                        {
                            columnValue.Add(cRng.Value.ToString());
                        }
                    }
                } while (result == null);

            }

            List<string> list = columnValue.Distinct().ToList();
            foreach (var item in list)
            {
                ExcelComps x = new ExcelComps();
                x.ExcelComp = item.ToString();
                lstXcelComps.Add(x); ;
            }
            return lstXcelComps;
        }

        public ObservableCollection<RIPLComps> GetRIPLComps()
        {
            ObservableCollection<RIPLComps> lstComps = new ObservableCollection<RIPLComps>();
            try
            {
                using (SqlCommand cmd = MyGlobalClass.OpenConnection())
                {
                    cmd.CommandText = String.Format("SELECT [Comp_Name],[Comp_ID] FROM [Import_78].[dbo].[Component] Where [Comp_Type_ID] = 1");
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            RIPLComps item = new RIPLComps();
                            item.RIPLCompName = reader["Comp_Name"].ToString();
                            item.CompID = Convert.ToInt32(reader["Comp_ID"]);

                            lstComps.Add(item);
                        }
                    }
                }
                return lstComps;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private void CompsOK_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                foreach (ExcelComps comps in ComponentMapping.Items)
                {
                    if (comps != null)
                    {
                        int x = 0;
                        var selectedcomp = comps.SelectedComp;//here you have selected item
                        var excelComp = comps.ExcelComp;
                        IEnumerable<RIPLComps> q1 = from lstComps in lstComps
                                                    where lstComps.RIPLCompName == selectedcomp.ToString()
                                                    select lstComps;
                        foreach (RIPLComps ma in q1)
                        {
                            x = Convert.ToInt32(ma.CompID);
                        }
                        origWindow.lstCompMapping.Add(new ComponentMappingClass()
                        {
                            compString = excelComp.ToString(),
                            compMapID = x
                        });
                    }
                   
                }
                
                this.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
