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
    public class Reference : PropertyChangedBase
    {
        [Mapping(ColumnName = "Reference Name")]
        private string _refname;
        public string Ref_Name
        {
            get { return _refname; }
            set
            {
                _refname = value;
                RaisePropertyChanged("RefName");
            }
        }
        [Mapping(ColumnName = "Reference ID")]
        private int _refID;
        public int Ref_ID
        {
            get { return _refID; }
            set
            {
                _refID = value;
                RaisePropertyChanged("Ref_ID");
            }
        }
        private string _selectedref;
        public string SelectedRef
        {
            get { return _selectedref; }
            set
            {
                _selectedref = value;
                RaisePropertyChanged("SelectedRef");
            }
        }


        //public Reference(string RefName, int Ref_ID)
        //{
        //    this.Ref_Name = Ref_Name;
        //    this.Ref_ID = Ref_ID;
        //}

    }
}
