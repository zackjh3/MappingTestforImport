using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MappingTest.Other_Classes
{
    public class SourceColumns
    {
        public string SourceCol { get; set; }

        public override string ToString()
        {
            return this.SourceCol;
        }
    }
}
