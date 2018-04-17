using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace MappingTest.RIPL_Classes
{
    public partial class RIPLVariables : PropertyChangedBase
    {

        public string VarName { get; set; }
        public override string ToString()
        {
            return this.VarName;
        }
        //private object _varname;

        //public object VarName
        //{
        //    get { return _varname; }
        //    set
        //    {
        //        _varname = value;
        //        RaisePropertyChanged("VarName");
        //    }
        //}
        private object _vartype;
        public object VarType
        {
            get { return _vartype; }
            set
            {
                _vartype = value;
                RaisePropertyChanged("VarType");
            }
        }
        private object _varID;
        public object VarID
        {
            get { return _varID; }
            set
            {
                _varID = value;
                RaisePropertyChanged("VarID");
            }
        }
        private string _selecteditem;
        public string SelectedItem
        {
            get { return _selecteditem; }
            set
            {
                _selecteditem = value;
                RaisePropertyChanged("SelectedItem");
            }
        }
    }
}
