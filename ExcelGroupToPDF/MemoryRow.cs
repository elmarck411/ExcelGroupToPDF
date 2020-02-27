using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGroupToPDF
{
    public class MemoryRow
    {
        public string CustNo        { get; set; }
        public string ShipTo        { get; set; }
        public string PONO          { get; set; }
        public string CSR           { get; set; }
        public string SLSNAME       { get; set; }
        public string OrderNo       { get; set; }
        public string ReleaseNo     { get; set; }
        public string OrdDate       { get; set; }
        public string PromiseDate   { get; set; }
        public string ItemNo        { get; set; }
        public string CustItemNo    { get; set; }
        public string CustDescrip   { get; set; }
        public string WHSE          { get; set; }
        public string OrdQty        { get; set; }
        public string OrdAvailQty   { get; set; }
        public string HoldTerms { get; set; }


        public MemoryRow()
        {

        }

    }
}
