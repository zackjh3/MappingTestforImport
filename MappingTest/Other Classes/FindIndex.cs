using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MappingTest.Other_Classes
{
    public class FindIndex
    {
        String _comp;
        public FindIndex(string comp)
        {
            _comp = comp;
        }
        public bool Equal(Components com)
        {
            return com.component.Equals(_comp, StringComparison.InvariantCultureIgnoreCase);
        }
    }
}
