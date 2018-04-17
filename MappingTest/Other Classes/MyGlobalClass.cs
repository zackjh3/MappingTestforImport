using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace MappingTest.Other_Classes
{
    public static class MyGlobalClass
    {
        #region SQL Connection
        public static string Sql()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = "localhost";
            builder.UserID = "zach.hine";              // update me
            builder.IntegratedSecurity = true;
            builder.Password = "password2";      // update me
            builder.InitialCatalog = "Import_78";
            return builder.ConnectionString;
        }
        public static SqlCommand OpenConnection()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            SqlConnection connection = new SqlConnection(Sql());
            cmd.Connection = connection;
            connection.Open();
            return cmd;
        }
        #endregion
    }
}
