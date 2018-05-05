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
using System.Windows.Navigation;
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
using MappingTest.Other_Classes;
using MappingTest;
using MappingTest.RIPL_Classes;

namespace MappingTest
{
    /// <summary>
    /// Interaction logic for AttMapping.xaml
    /// </summary>
    public partial class AttMapping : Window
    {

        public static List<ModelAtt> NewSQLAtt { get; set; }
        public static List<MyModel> MyDataGridItems { get; set; }
        public static DataTable myAtts { get; set; }
        public int v = 0;
        MainWindow originalWindow;

        public AttMapping(MainWindow incomingWindow)
        {
            
            InitializeComponent();
            originalWindow = incomingWindow;
            AttributeMapping.AttMapp();
            NewSQLAtt = new List<ModelAtt>();
            NewSQLAtt = AttributeMapping.SQLAtt;
            MyDataGridItems = new List<MyModel>();
            
            foreach (DataRow item in AttributeMapping.myAtts.Rows)
            {
                MyDataGridItems.Add(new MyModel() { XcelAtt = item[0].ToString() });
            }


            IEnumerable<RIPLVariables> result = from s in originalWindow.VarDataGridItems
                                                where s.VarName == MainWindow.ExcelMappedVar
                                                select s;

            foreach (RIPLVariables rv in result)
            {
                v = Convert.ToInt32(rv.VarID);
            }


            // originalWindow.VarDataGridItems.Where(z => z.VarName == MainWindow.ExcelMappedVar).Select(x => x.VarID));
            // MessageBox.Show(hi.ToString());


        }
        
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            
        }
        
        public void OK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (MyModel model in AttMap.Items)
                {
                    int x = 0;
                    var selecteditem = model.SelectedItem;//here you have selected item
                    var excelAtt = model.XcelAtt;
                    IEnumerable<ModelAtt> q1 = from SQLAtt in AttributeMapping.SQLAtt
                                               where SQLAtt.modelatt == selecteditem.ToString()
                                               select SQLAtt;
                    foreach (ModelAtt ma in q1)
                    {
                        x = Convert.ToInt32(ma.attID);
                    }
                    originalWindow.passList.Add(new AttMapList()
                    {
                        attString = excelAtt.ToString(),
                        attID = x,
                        VarID = v

                    });
                 
                }
               
                
                //mappingWin.PassList(this.passList);
                this.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
    }
}
//public void Take1DArray(object arr)
//{
//    System.Array column = (Array)arr;

//    List<string> list = column.OfType<string>().ToList();

//    column = list.Distinct<string>().ToArray();

//    foreach (var item in column)
//    {
//        MessageBox.Show(item.ToString());
//    }
//}
