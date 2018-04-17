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

namespace MappingTest
{
    /// <summary>
    /// Interaction logic for AttMapping.xaml
    /// </summary>
    public partial class AttMapping : Window
    {
        public List<MyModel> MyDataGridItems { get; set; }
        public List<ModelAtt> SQLAtt { get; set; }
        public DataTable myAtts { get; set; }
        public string selsheet = "HCA";
        MainWindow originalWindow;

        public AttMapping(MainWindow incomingWindow)
        {
            
            InitializeComponent();
            originalWindow = incomingWindow;
            SQLAtt = new List<ModelAtt>();
            SQLAtt = GetAttributes(MainWindow.RIPLMappedVar);

            MyDataGridItems = new List<MyModel>();

            myAtts = DemoDistinct();
            foreach (DataRow item in myAtts.Rows)
            {
                MyDataGridItems.Add(new MyModel() { XcelAtt = item[0].ToString() });
            }
           
            AttMap.ItemsSource = MyDataGridItems;
            
          
         
        }
        public string Sql()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = "localhost";
            builder.UserID = "zach.hine";              // update me
            builder.IntegratedSecurity = true;
            builder.Password = "password2";      // update me
            builder.InitialCatalog = "master";
            return builder.ConnectionString;
        }
        public SqlCommand OpenConnection()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            SqlConnection connection = new SqlConnection(Sql());
            cmd.Connection = connection;
            connection.Open();
            return cmd;
        }
        public List<ModelAtt> GetAttributes(string variable)
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
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            
        }
        public void Take1DArray(object arr)
        {
            System.Array column = (Array)arr;

            List<string> list = column.OfType<string>().ToList();

            column = list.Distinct<string>().ToArray();

            foreach (var item in column)
            {
                MessageBox.Show(item.ToString());
            }
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
                    IEnumerable<ModelAtt> q1 = from SQLAtt in SQLAtt
                                               where SQLAtt.modelatt == selecteditem.ToString()
                                               select SQLAtt;
                    foreach (ModelAtt ma in q1)
                    {
                        x = Convert.ToInt32(ma.attID);
                    }
                    originalWindow.passList.Add(new AttMapList()
                    {
                        attString = excelAtt.ToString(),
                        attID = x
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


        public class ModelAtt
        {
            public string modelatt { get; set; }
            public override string ToString()
            {
                return this.modelatt;
            }
            public int attID { get; set; }
        }

        private string GetConnectionString()
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

        private DataTable DemoDistinct()
        {
           
            List<string> dateList = new List<string>();
            DataTable dt = new DataTable();

            string connectionString = GetConnectionString();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                cmd.CommandText = String.Format("SELECT DISTINCT ["+ MainWindow.ExcelMappedVar+"] FROM [" + selsheet +"$" + "] WHERE [" + MainWindow.ExcelMappedVar + "] IS NOT NULL");





                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
            }

            return dt;
        }
    }
    
    public class MyModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private object _xcelatt;
        public object XcelAtt
        {
            get { return _xcelatt; }
            set
            {
                _xcelatt = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("XcelAtt"));
            }
        }

        private object _selectedItem;
        public object SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("SelectedItem"));
            }
        }
        public virtual void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, e);
        }
    }
}

