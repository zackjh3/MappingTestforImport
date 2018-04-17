using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MappingTest.RIPL_Classes
{
    public class RIPLComps
    {
        public string RIPLCompName { get; set; }
        public override string ToString()
        {
            return this.RIPLCompName;
        }
        public int CompID { get; set; }

    }
}
