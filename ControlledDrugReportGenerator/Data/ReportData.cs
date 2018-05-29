using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ControlledDrugReportGenerator.Data
{
    class ReportData
    {
        public string PatientID { get; set; }
        public string PatientName { get; set; }
        public string TransactionDate { get; set; }
        public string OrderID { get; set; }
        public string OrderStartTime { get; set; }
        public string Dose { get; set; }
        public string UserName { get; set; }
        public string Quantity { get; set; }
        public string QuantityUnit { get; set; }
        public string EndDose { get; set; }
        public string UsingUnit { get; set; }
        public string MedName { get; set; }
        public string MedID { get; set; }
    }
}
