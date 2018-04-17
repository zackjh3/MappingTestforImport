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

namespace MappingTest.Other_Classes
{
    public class Reference
    {
        [Mapping(ColumnName = "Reference Name")]
        public string Ref_Name { get; set; }
        [Mapping(ColumnName = "Reference ID")]
        public int Ref_ID { get; set; }

        //public Reference(string RefName, int Ref_ID)
        //{
        //    this.Ref_Name = Ref_Name;
        //    this.Ref_ID = Ref_ID;
        //}
        public static List<Reference> GetReferences()
        {
            List<Reference> refs = new List<Reference>();
            try
            {
                using (SqlCommand cmd = MyGlobalClass.OpenConnection())
                {
                    cmd.CommandText = "SELECT [Ref_Name],[Ref_ID] FROM [Import_78].[dbo].[Ref_Def]";
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Reference item = new Reference();
                            item.Ref_Name = reader["Ref_Name"].ToString();
                            item.Ref_ID = Convert.ToInt32(reader["Ref_ID"]);

                            refs.Add(item);
                        }
                    }
                }
                return refs;

            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
