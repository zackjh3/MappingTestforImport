using System;
using System.Collections.Generic;
using System.Linq;
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

    public class ModelAtt : PropertyChangedBase
    {
        public string modelatt { get; set; }
        public override string ToString()
        {
            return this.modelatt;
        }
        private string _selecteditem;
        public string SelectedSqlAtt
        {
            get { return _selecteditem; }
            set
            {
                _selecteditem = value;
                RaisePropertyChanged("SelectedSqlAtt");
            }
        }
        public static SqlCommand OpenConnection()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            SqlConnection connection = new SqlConnection(MyGlobalClass.Sql());
            cmd.Connection = connection;
            connection.Open();
            return cmd;
        }
        public int attID { get; set; }
        public static List<ModelAtt> GetAttributes(string variable)
        {

            List<ModelAtt> dtAttributes = new List<ModelAtt>();
            try
            {
                using (SqlCommand cmd = OpenConnection())
                {
                    cmd.CommandText = String.Format("SELECT [Description],[Att_ID] FROM[Import_78].[dbo].[Attributes]  WHERE[Import_78].[dbo].[Attributes].Att_ID IN(SELECT[Import_78].[dbo].[Att_Link].Att_ID FROM[Import_78].[dbo].[Att_Link] WHERE[Import_78].[dbo].[Att_Link].Var_ID IN(Select[Import_78].[dbo].[Variables].Var_ID FROM[Import_78].[dbo].[Variables] WHERE[Import_78].[dbo].[Variables].Var_Description = '{0}'))", variable);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ModelAtt item = new ModelAtt();
                            item.modelatt = reader["Description"].ToString();
                            item.attID = Convert.ToInt32(reader["Att_ID"]);

                            dtAttributes.Add(item);
                        }

                    }
                }

                return dtAttributes;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
    }
}
