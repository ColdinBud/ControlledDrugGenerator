using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ControlledDrugReportGenerator.Data
{
    class LabelData
    {
        public string Device { get; set; }
        public string Drawer { get; set; }
        public string MedID { get; set; }
        public string MedName { get; set; }
        public int Min { get; set; }
        public int Max { get; set; }
        public int Current { get; set; }
        public int Amount { get; set; }
    }
}
