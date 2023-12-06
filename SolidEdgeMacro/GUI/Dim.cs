using SolidEdgeFrameworkSupport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GUI
{
    internal class Dim
    {
        public bool IsSelected { get; set; }
        public string Name { get; set; }
        public double Value { get; set; }
        public Dimension Dimension { get; set; }
    }
}
