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
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Common;


namespace MappingTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    
    public partial class MainWindow : Window
    {
        public string HELP;
        public static string selectedSheet = "";
        public static string ExcelMappedVar = "";
        public static string RIPLMappedVar = "";
        public List<RIPLVariables> VarDataGridItems { get; set; }
        public ObservableCollection<ExcelVariables> ExcelVar { get; set; }
        public static Excel.Worksheet excelSheet { get; set; }
        public ObservableCollection<string> Types { get; set; }
        public ObservableCollection<InputModelColumns> InputModelCol { get; set; }
        public static ObservableCollection<SourceColumns> SourceCol { get; set; }
        public static string selectedFile = "C:\\Users\\zach.hine\\American Innovations\\Import\\TestData.xlsx";
        private string transform = string.Empty;
        
        public List<AttMapList> passList = new List<AttMapList>();
        public List<ComponentMappingClass> lstCompMapping = new List<ComponentMappingClass>();
        public static List<VarMapList> VMapList = new List<VarMapList>();

        public MainWindow()
        {
            InitializeComponent();
            HELP = "HELP";
        }
        public void MappingWindowLoaded(object sender, RoutedEventArgs e)
        {

            //string selectedFile = "C:\\Users\\zach.hine\\American Innovations\\Import\\TestData.xlsx";
            //string selectedFile = ((MainWindow)Application.Current.MainWindow).filePath.Text;
            var results = GetAllWorksheets(selectedFile);

            Selected_Models.Items.Add("Haz Liq - High Consequence Areas");
            string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + selectedFile + ";Extended Properties=Excel 12.0;");
            BuildReferenceList();
            List<string> inputList = new List<string>();



            // string selectedFile = ((MainWindow)Application.Current.MainWindow).filePath.Text;

            List<string> worksheets = new List<string>();
            worksheets = GetWorksheets(selectedFile);
            cbSourceSheet.ItemsSource = worksheets;

        }
        public static Sheets GetAllWorksheets(string fileName)
        {
            Sheets theSheets = null;

            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                theSheets = wbPart.Workbook.Sheets;
                document.Close();
            }

            return theSheets;
        }
        public List<ExcelVariables> GetColumns(string fileName, string sheet)
        {
            //DataTable dtColumns = new DataTable();
            //dtColumns.Columns.Add("Columns", typeof(string));
            List<ExcelVariables> lstXcelVar = new List<ExcelVariables>();
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

            excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(num);
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
        public void Selected_Models_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbSourceSheet.IsEnabled = true;
            int model_ID;
            string selModel = Selected_Models.SelectedItem.ToString();
            using (SqlConnection connection = new SqlConnection(MyGlobalClass.Sql()))
            {

                connection.Open();
                string sqlString = String.Format("SELECT Model_ID FROM [Import_78].[dbo].[Model] WHERE [Model_Name] = '{0}'", selModel);
                SqlCommand cmd = new SqlCommand(sqlString, connection);
                model_ID = Convert.ToInt32(cmd.ExecuteScalar());
                SelectedInputModels a = new SelectedInputModels(selModel, model_ID);

            }
        }
        public int GetModelPositionType(string model)
        {
            try
            {
                int positionType = 0;
                using (SqlCommand cmd = MyGlobalClass.OpenConnection())
                {
                    cmd.CommandText = String.Format("Select [Positioning] From [Import_78].[dbo].[Model] Where Model_Name = '{0}'", model);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {

                        while (reader.Read())
                        {
                            positionType = Convert.ToInt32(reader["Positioning"]);

                        }
                    }
                }
                return positionType;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public string GetStationingNames(int varID)
        {
            try
            {
                string station = "";
                using (SqlCommand cmd = MyGlobalClass.OpenConnection())
                {
                    cmd.CommandText = String.Format("Select [Var_Description] From [Import_78].[dbo].[Variables] Where Var_ID = '{0}'", varID);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {

                        while (reader.Read())
                        {
                            station = Convert.ToString(reader["Var_Description"]);

                        }
                    }
                }
                return station;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        private void BuildReferenceList()
        {
            cbReferences.Items.Clear();
            try
            {
                List<Reference> refs = Reference.GetReferences();
                #region Add Feature Classes to combobox
                int idx = 0;
                cbReferences.Items.Add("");
                for (int i = 0; i < refs.Count; i++)
                {
                    cbReferences.Items.Add(refs[i].Ref_Name);
                }

                cbReferences.SelectedIndex = idx > 0 ? idx : 0;

                #endregion Feature Classes to combobox

            }
            catch (Exception e)
            {
                throw e;
            }


        }
        private void cbSourceSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            var SourceSheet = sender as ComboBox;
            selectedSheet = SourceSheet.SelectedItem as string;
            
            VarDataGridItems = GetVariables(Selected_Models.SelectedItem.ToString());

            int z = GetModelPositionType(Selected_Models.SelectedItem.ToString());

            RIPLVariables addRIPL = new RIPLVariables();
            addRIPL.VarName = "Component";
            addRIPL.VarID = 100;
            addRIPL.VarType = 16;
            VarDataGridItems.Add(addRIPL);

            RIPLVariables addRIPL2 = new RIPLVariables();

            if (z == 5)
            {
                addRIPL2.VarName = GetStationingNames(1);
                addRIPL2.VarID = 1;
                addRIPL2.VarType = 18;
                VarDataGridItems.Add(addRIPL2);
                RIPLVariables item3 = new RIPLVariables();
                item3.VarName = GetStationingNames(4);
                item3.VarID = 4;
                item3.VarType = 18;
                VarDataGridItems.Add(item3);
            }
            else if (z == 2)
            {
                addRIPL2.VarName = GetStationingNames(2);
                addRIPL2.VarID = 1;
                addRIPL2.VarType = 18;
                VarDataGridItems.Add(addRIPL2);
            }
            VarMapping.ItemsSource = VarDataGridItems;


            List<InputModelColumns> inputmodelcol = new List<InputModelColumns>();


            //string selectedFile = ((MainWindow)Application.Current.MainWindow).filePath.Text;

            ExcelVar = new ObservableCollection<ExcelVariables>();
            ExcelVar = ExcelVariables.GetColumns(selectedFile, selectedSheet);

        }
        
        public List<RIPLVariables> GetVariables(string model)
        {
            List<RIPLVariables> lstVariables = new List<RIPLVariables>();
            try
            {
                using (SqlCommand cmd = MyGlobalClass.OpenConnection())
                {
                    cmd.CommandText = String.Format("SELECT [Var_Description],[Var_ID],[Var_Type] From [Import_78].[dbo].[Variables] WHERE [Import_78].[dbo].[Variables].[Var_ID] IN (SELECT [Import_78].[dbo].[Model_Link].[Var_ID] FROM [Import_78].[dbo].[Model_Link] WHERE [Import_78].[dbo].[Model_Link].[Model_ID] IN (SELECT [Import_78].[dbo].[Model].[Model_ID] FROM [Import_78].[dbo].[Model] WHERE [Model_Name] = '{0}'))", model);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            RIPLVariables item = new RIPLVariables();
                            item.VarName = reader["Var_Description"].ToString();
                            item.VarID = Convert.ToInt32(reader["Var_ID"]);
                            item.VarType = Convert.ToInt32(reader["Var_Type"]);

                            lstVariables.Add(item);
                        }

                    }
                }
                return lstVariables;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        private void AttMapp_Click(object sender, RoutedEventArgs e)
        {
            
            int index = VarMapping.SelectedIndex;
            var model = VarMapping.Items[index] as RIPLVariables;
            var selecteditem = model.SelectedItem;
            ExcelMappedVar = Convert.ToString(selecteditem);
            var x = VarDataGridItems[index];
            RIPLMappedVar = Convert.ToString(x);
      


            AttMapping AttMappingWindow = new AttMapping(this);
            AttMappingWindow.Show();
        }
        public List<string> GetWorksheets(string fileName)
        {
            int num = 1;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(fileName.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(num);
            Excel.Range excelRange = excelSheet.UsedRange;
            List<string> list = new List<string>();
            foreach (Excel.Worksheet exSheet in excelBook.Sheets)
            {
                list.Add(exSheet.Name);
            }
            return list;
        }

        private int FindRowIndex(DataGridRow row)
        {
            DataGrid dataGrid =
                ItemsControl.ItemsControlFromItemContainer(row)
                as DataGrid;

            int index = dataGrid.ItemContainerGenerator.
                IndexFromContainer(row);

            return index;
        }

        private void CompMapping_Click(object sender, RoutedEventArgs e)
        {
                //foreach (RIPLVariables vari in VarMapping.Items)
                //{
                    
                //    string y="";
                   
                //    var selecteditem = vari.SelectedItem;//here you have selected item
                //    var RIPLVar = vari.VarName;
                //if (selecteditem != null)
                //{
                //    IEnumerable<ExcelVariables> q1 = from ExcelVar in ExcelVar
                //                                     where ExcelVar.XcelVar == selecteditem.ToString()
                //                                     select ExcelVar;
                //    foreach (ExcelVariables ma in q1)
                //    {
                //        y = Convert.ToString(ma.XcelVar);
                //    }
                //    VMapList.Add(new VarMapList()
                //    {
                //        RIPLVarID = Convert.ToInt32(vari.VarID),
                //        ExcelVarString = y

                //    });
                //}
               
               
                //}

       

            CompMapping CompMappingWindow = new CompMapping(this);
            CompMappingWindow.Show();
            MessageBox.Show(lstCompMapping.Count.ToString());
        }

        //public static void ImportTableToSQLServer(string connStrSource, string sourceTable, string connStrDestination)
        //{
        //    //Create temp table for data
        //    string tempTable = "ImportTempTable";
        //    // Create database connections
        //    //string connStrSource = "Data Source= Import78
        //    SqlConnection sqlConnDestination = new SqlConnection(MyGlobalClass.Sql());

        //    if (sqlConnDestination.State == System.Data.ConnectionState.Broken
        //      || sqlConnDestination.State == System.Data.ConnectionState.Closed)
        //    { sqlConnDestination.Open(); }

        //    // Create the destination database object
        //    ServerConnection connDestination = new ServerConnection(sqlConnDestination);
        //    Server server = new Server(connDestination);
        //    Database destinationDB = server.Databases[sqlConnDestination.Database];

        //    Microsoft.SqlServer.Management.Smo.Table copiedtable = null;
        //    DocumentFormat.OpenXml.Spreadsheet.Table sTable = ;
        //    if (sTable != null)
        //    {
        //        copiedtable = new Microsoft.SqlServer.Management.Smo.Table(destinationDB, tempTable, sTable.Schema);

        //        // Create the new destination table
        //        CreateColumnsFromExceltable(ref sTable, ref copiedtable);
        //        copiedtable.AnsiNullsStatus = sTable.AnsiNullsStatus;
        //        copiedtable.QuotedIdentifierStatus = sTable.QuotedIdentifierStatus;
        //        copiedtable.TextFileGroup = sTable.TextFileGroup;
        //        copiedtable.FileGroup = sTable.FileGroup;
        //        copiedtable.Create();
        //    }
        //    else
        //    {

        //    }
        //}
      //  private static void CreateColumnsFromExceltable(ref Microsoft.SqlServer.Management.Smo.Table sourceTable, ref Microsoft.SqlServer.Management.Smo.Table copiedTable)
      //  {

      //      try
      //      {
      //          // Create a source SQL Server object
      //          var serverSource = sourceTable.Parent.Parent;
      //          var serverDestination = copiedTable.Parent.Parent;

      //          // Re-create each source table column in the new destination table
      //          foreach (DocumentFormat.OpenXml.SpreadsheetColumn source in sourceTable.Columns)
      //          {
      //              var column = new Column(copiedTable, source.Name, source.DataType);
      //              column.Collation = source.Collation;
      //              column.Nullable = source.Nullable;
      //              column.Computed = source.Computed;
      //              column.ComputedText = source.ComputedText;
      //              column.Default = source.Default;

      //              if (source.DefaultConstraint != null)
      //              {
      //                  column.AddDefaultConstraint(copiedTable.Name + SEPERATOR + source.DefaultConstraint.Name);
      //                  column.DefaultConstraint.Text = source.DefaultConstraint.Text;
      //              }

      //              column.IsPersisted = source.IsPersisted;
      //              column.DefaultSchema = source.DefaultSchema;
      //              column.RowGuidCol = source.RowGuidCol;

      //              if (serverSource.VersionMajor >= 10 && serverDestination.VersionMajor >= 10)
      //              {
      //                  column.IsFileStream = source.IsFileStream;
      //                  column.IsSparse = source.IsSparse;
      //                  column.IsColumnSet = source.IsColumnSet;
      //              }

      //              copiedTable.Columns.Add(column);
      //          }
      //      }
      //      catch (Exception e)
      //      {
      //          Logger.LogError("CreateColumnsFromSQLtable()", e);
      //          throw;
      //      }
      //  }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            foreach (RIPLVariables vari in VarMapping.Items)
            {
                string y = "";
                var selecteditem = vari.SelectedItem;//here you have selected item
                var RIPLVar = vari.VarName;
                if (selecteditem != null)
                {
                    IEnumerable<ExcelVariables> q1 = from ExcelVar in ExcelVar
                                                     where ExcelVar.XcelVar == selecteditem.ToString()
                                                     select ExcelVar;
                    foreach (ExcelVariables ma in q1)
                    {
                        y = Convert.ToString(ma.XcelVar);
                    }
                    VMapList.Add(new VarMapList()
                    {
                        RIPLVarID = Convert.ToInt32(vari.VarID),
                        ExcelVarString = y

                    });
                }


            }

            string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + selectedFile + ";Extended Properties=Excel 12.0;");
            string tempTable = "AImportTempTable";
            string fileType = ".xlsx";
            selectedSheet = cbSourceSheet.SelectedItem as string;
            
            ImportTableFromFile(fileType, connectionString, selectedSheet, MyGlobalClass.Sql(), tempTable);

            using (SqlConnection connection = new SqlConnection(MyGlobalClass.Sql()))
            {
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    connection.Open();
                    foreach (var item in VMapList)
                    {
                        if (passList != null)
                        {
                            foreach (var item2 in passList)
                            {
                                string sqlString = String.Format("UPDATE AImportTempTable SET [{2}]={0} WHERE [{2}]='{1}'", item2.attID, item2.attString, item.ExcelVarString);
                                cmd.CommandText = sqlString;
                                cmd.ExecuteNonQuery();
                            }
                            
                        }
                        string sqlcmd = String.Format("EXEC sp_RENAME 'AImportTempTable.{0}', 'var{1}', 'COLUMN'", item.ExcelVarString, item.RIPLVarID);
                        cmd.CommandText = sqlcmd;
                        cmd.ExecuteNonQuery();
                    }

                    //foreach (var item2 in passList)
                    //{
                    //    string sqlString = String.Format("UPDATE AImportTempTable SET [HCA Type]={0} WHERE [HCA Type]='{1}'", item2.attID, item2.attString);
                    //    cmd.CommandText = sqlString;
                    //    cmd.ExecuteNonQuery();
                    //}

                    foreach (var item in lstCompMapping)
                    {
                        string sqlString2 = String.Format("UPDATE AImportTempTable SET [Component] = {0} WHERE [Component]='{1}'", item.compMapID, item.compString);
                        cmd.CommandText = sqlString2;
                        cmd.ExecuteNonQuery();
                    }
                    connection.Close();
                    MessageBox.Show("Import Complete");
                }
             
               
            }

        }
   








        /// <summary>
        /// Imports a table of data from a file
        /// </summary>
        /// <param name="fileType">Type of file</param>
        /// <param name="connStrSource">SQL Connection string to the source database</param>
        /// <param name="sourceTable">Source table name</param>
        /// <param name="connStrDestination">SQL Connection string to the destination database</param>
        /// <param name="destinationTable">Destination table name</param>
        public static void ImportTableFromFile(string fileType, string connStrSource, string sourceTable, string connStrDestination, string destinationTable)
        {
           

            string errMsg = "Create OleDB Connection object";

            try
            {
                // Create database connections
                OleDbConnection sourceConn = new OleDbConnection(connStrSource);

                errMsg = "Create SQL Connection object";

                //string connStrDestination = "Data Source=Forsberg\\SQL2008;Initial Catalog=RIPL_v7_2010;Integrated Security=True;";
                SqlConnection sqlConnDestination = new SqlConnection(connStrDestination);

                // Open connections if they are closed
                if (sourceConn.State == System.Data.ConnectionState.Broken
                    || sourceConn.State == System.Data.ConnectionState.Closed)
                { sourceConn.Open(); }

                errMsg = "Source File OleDB Connection Opened: " + connStrSource;

                if (sqlConnDestination.State == System.Data.ConnectionState.Broken
                    || sqlConnDestination.State == System.Data.ConnectionState.Closed)
                { sqlConnDestination.Open(); }

                errMsg = "SQL Connection Opened: " + connStrDestination;

                // Create the destination database object
                ServerConnection connDestination = new ServerConnection(sqlConnDestination);
                Server server = new Server(connDestination);

                Database destinationDB = server.Databases[sqlConnDestination.Database];

                errMsg = "Get the source table schema";

                // Get the source data schema
                OleDbCommand cmd = new OleDbCommand("select * from [" + sourceTable + "$" + "]", sourceConn);
                var reader = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
                DataTable sTable = reader.GetSchemaTable();

                // Delete the destination table if it exists
                if (((Microsoft.SqlServer.Management.Smo.SmoObjectBase)(destinationDB.Tables[destinationTable])) != null)
                {
                    if (destinationDB.Tables[destinationTable].State == SqlSmoState.Existing)
                    { destinationDB.Tables[destinationTable].Drop(); }
                }

                errMsg = "Create the destination table";

                // Create a new empty destination table
                Microsoft.SqlServer.Management.Smo.Table copiedtable = new Microsoft.SqlServer.Management.Smo.Table(destinationDB, destinationTable);
                CreateColumnsFromFileTable(ref sTable, ref copiedtable);
                copiedtable.Create();

                errMsg = "";

                // Import the data from the source to the new destination table
                GetSourceDataFromFile(ref fileType, ref destinationDB, ref sourceConn, ref sourceTable, ref sqlConnDestination, ref destinationTable);
            }
            catch (Exception ex)
            {
               

                if (errMsg != "")
                {
                    ex.HelpLink = errMsg;
                }

                throw ex;
            }
        }

        /// <summary>
        /// Imports data from the source file to a SQL Server table
        /// </summary>
        /// <param name="fileType">Type of file</param>
        /// <param name="destinationDB">Destination SQL Server Database object</param>
        /// <param name="sourceConn">OleDb Connection to the source file</param>
        /// <param name="sourceTable">Source table name</param>
        /// <param name="sqlConnDestination">SQL Connection to the destination database</param>
        /// <param name="destinationTable">Destination table name</param>
        private static void GetSourceDataFromFile(ref string fileType, ref Database destinationDB, ref OleDbConnection sourceConn,
            ref string sourceTable, ref SqlConnection sqlConnDestination, ref string destinationTable)
        {
           

            string errMsg2 = "Create a SMO destination table: " + destinationTable;

            try
            {
                // Create the destination table object
                Microsoft.SqlServer.Management.Smo.Table dTable = destinationDB.Tables[destinationTable];

                errMsg2 = "Create the SqlBulkCopy object";

                // Create the destination buld copy object
                SqlBulkCopy bcp = new SqlBulkCopy(sqlConnDestination);
                bcp.DestinationTableName = "[" + destinationTable + "]";
                bcp.BatchSize = 100000;
                bcp.BulkCopyTimeout = 300;

                // Get the list of columns
                string columnNames = "*";
                if (fileType != "PCS")
                {
                    columnNames = GetColumnNamesFromTable(ref dTable, ref bcp);
                }

                errMsg2 = "File Type: " + fileType + ", Columns: " + columnNames;

                // Create the source command object
                System.Data.OleDb.OleDbCommand cmdSource = sourceConn.CreateCommand();
                cmdSource.CommandText = "SELECT " + columnNames + " FROM [" + sourceTable + "$" + "] s";
                cmdSource.CommandType = CommandType.Text;

                errMsg2 = "Create the source OleDB command object: " + cmdSource.CommandText;

                // Import the data
                System.Data.OleDb.OleDbDataReader reader = cmdSource.ExecuteReader();

                errMsg2 = "Copy the source data to the SQL Server Import_ table";

                bcp.WriteToServer(reader);

                errMsg2 = "Data Copied.";
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Copies the columns from the source table to the new table.
        /// </summary>
        /// <param name="sourcetable">Source file table</param>
        /// <param name="copiedtable">SQL Server destination table</param>
        private static void CreateColumnsFromFileTable(ref DataTable sTable, ref Microsoft.SqlServer.Management.Smo.Table copiedtable)
        {
            

            try
            {
                bool addCol = false;
                var nameCol = sTable.Columns["ColumnName"];

                // Re-create each source table column in the new destination table
                foreach (DataRow row in sTable.Rows)
                {
                    //Console.WriteLine(row[nameCol] + ": " + row[5].ToString());

                    Microsoft.SqlServer.Management.Smo.Column newCol = new Microsoft.SqlServer.Management.Smo.Column(copiedtable, row[nameCol].ToString());
                    addCol = true;

                    switch (row[5].ToString())
                    {
                        case "System.Byte[]":
                            addCol = false;
                            break;

                        case "System.Byte":
                            newCol.DataType = DataType.TinyInt;
                            break;

                        case "System.Boolean":
                        case "System.Int16":
                            newCol.DataType = DataType.SmallInt;
                            break;

                        case "System.Int":
                        case "System.Int32":
                            newCol.DataType = DataType.Int;
                            break;

                        case "System.Single":
                            newCol.DataType = DataType.Real;
                            break;

                        case "System.Double":
                        case "System.Decimal":
                            newCol.DataType = DataType.Float;
                            break;

                        case "System.DateTime":
                            newCol.DataType = DataType.DateTime;
                            break;

                        default:
                            if (row[2].ToString().Length > 0 && Convert.ToInt32(row[2]) > 0)
                            {
                                if (Convert.ToInt32(row[2]) > 255)
                                {
                                    newCol.DataType = DataType.VarChar(255);
                                }
                                else
                                {
                                    newCol.DataType = DataType.VarChar(Convert.ToInt32(row[2]));
                                }
                            }
                            else
                            {
                                newCol.DataType = DataType.VarChar(255);
                            }

                            break;

                    }

                    if (addCol)
                    {
                        newCol.Nullable = true;

                        copiedtable.Columns.Add(newCol);
                    }
                }
            }
            catch (Exception e)
            {
                
            }
        }
        private static string GetColumnNamesFromTable(ref Microsoft.SqlServer.Management.Smo.Table destinationTable, ref SqlBulkCopy bcp)
        {
           

            string colNames = "";

            try
            {
                // Create the column name list and configure the Bulk Copy column mappings
                foreach (Microsoft.SqlServer.Management.Smo.Column source in destinationTable.Columns)
                {
                    SqlBulkCopyColumnMapping sqlbccm = new SqlBulkCopyColumnMapping(source.Name, source.Name);
                    bcp.ColumnMappings.Add(sqlbccm);

                    if (colNames.Length > 0)
                    {
                        colNames += ", s.[" + source.Name + "]";
                    }
                    else
                    {
                        colNames = "s.[" + source.Name + "]";
                    }
                }
            }
            catch (Exception e)
            {
              
                throw e;
            }
            return colNames;
        }


    }
}
    

