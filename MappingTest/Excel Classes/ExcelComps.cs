using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MappingTest.Other_Classes
{
    public class ExcelComps : PropertyChangedBase
    {
        public string ExcelComp { get; set; }
        public override string ToString()
        {
            return this.ExcelComp;
        }
        private string _selectedcomp;
        public string SelectedComp
        {
            get { return _selectedcomp; }
            set
            {
                _selectedcomp = value;
                RaisePropertyChanged("SelectedComp");
            }
        }
    }
}
