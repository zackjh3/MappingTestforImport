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

namespace MappingTest.Other_Classes
{
    public partial class AttributeMapping
    {
        public static List<ModelAtt> SQLAtt { get; set; }
        public static List<MyModel> MyDataGridItems { get; set; }
        public static DataTable myAtts { get; set; }

        public static void AttMapp()
        {
            SQLAtt = new List<ModelAtt>();
            SQLAtt = ModelAtt.GetAttributes(MainWindow.RIPLMappedVar);


            myAtts = DemoDistinct(MainWindow.ExcelMappedVar);
            
  
            
        }

        public static string selsheet = "Pipe Segment";
        public static string GetConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = MainWindow.selectedFile;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }
        public static DataTable DemoDistinct(string var)
        {

            List<string> dateList = new List<string>();
            DataTable dt = new DataTable();

            string connectionString = GetConnectionString();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                cmd.CommandText = String.Format("SELECT DISTINCT [" + var + "] FROM [" + selsheet + "$" + "] WHERE [" + var + "] IS NOT NULL");
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);

                //cmd.CommandText = String.Format("SELECT [VarID_]WHERE [" + MainWindow.ExcelMappedVar + "] IS NOT NULL");
            }

            return dt;
        }
    }
}
