using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MappingTest.Other_Classes
{
    public class Transform
    {
        private string _transform;

        public string _Transform
        {
            get
            {
                return this._transform;
            }

            set
            {
                this._transform = value;
            }
        }

        public override string ToString()
        {
            return this._transform;
        }
    }
}
