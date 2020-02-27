using System;
using Microsoft.Office.Interop;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelGroupToPDF
{
    class Program
    {
        public const int NumberOfColumns = 16;

        static void Main(string[] args)
        {

            //Open the File and spreadsheet
            var spreadsheetLocation = Path.Combine(Directory.GetCurrentDirectory(), "_ALP_OpenOrderCSR_TEST_NoParams.xls");
            var exApp = new Application();
            var exWbk = exApp.Workbooks.Open(spreadsheetLocation);
            Worksheet exWks = exWbk.Sheets["Sheet1"];
            
            Microsoft.Office.Interop.Excel.Range xlRange = (Microsoft.Office.Interop.Excel.Range)exWks.UsedRange as Microsoft.Office.Interop.Excel.Range;

            //xlRange.Group(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //xlRange.OutlineLevel = 1;
            exWks.EnableAutoFilter = true;

            // var memberValue = (string)(exWks.Cells[4, 5] as Microsoft.Office.Interop.Excel.Range).Value;
            //Microsoft.Office.Interop.Excel.Range groupedRange;
            List<MemoryRow> objs =  assignProperties(xlRange);

            var groupedFields = (from o in objs
                                 group o by o.CustNo);


            foreach (var custNoGroup in groupedFields)
            {

                Worksheet newWorkSheet = exWbk.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Range groupRange = newWorkSheet.get_Range("A1", "P" + custNoGroup.Count()  );
                

                string[,] arrayGroup = new string[custNoGroup.Count(), NumberOfColumns];

                int i = 0;
                foreach (var row in custNoGroup)
                {
                    arrayGroup[i, 0] =  row.CustNo;
                    arrayGroup[i, 1] =  row.ShipTo;
                    arrayGroup[i, 2] =  row.PONO;
                    arrayGroup[i, 3] =  row.CSR;
                    arrayGroup[i, 4] =  row.SLSNAME;
                    arrayGroup[i, 5] =  row.OrderNo;
                    arrayGroup[i, 6] =  row.ReleaseNo;
                    arrayGroup[i, 7] =  row.OrdDate;
                    arrayGroup[i, 8] =  row.PromiseDate;
                    arrayGroup[i, 9] = row.ItemNo;
                    arrayGroup[i, 10] = row.CustItemNo;
                    arrayGroup[i, 11] = row.CustDescrip;
                    arrayGroup[i, 12] = row.WHSE;
                    arrayGroup[i, 13] = row.OrdQty;
                    arrayGroup[i, 14] = row.OrdAvailQty;
                    arrayGroup[i, 15] = row.HoldTerms;
                    i++;
                  //  groupRange.Insert(Type.Missing, arrayGroup);
                }
                groupRange.Value = arrayGroup;

                //groupRange.ExportAsFixedFormat
                // Save into a PDF.
                string filename = "CustomerInfo.pdf";
                const int xlQualityStandard = 0;
                //exWks.ExportAsFixedFormat(
                groupRange.ExportAsFixedFormat(
                    Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                    filename, xlQualityStandard, true, true,
                    Type.Missing, Type.Missing, true, Type.Missing);
            }

            

           
                
        }

        public static List<MemoryRow>  assignProperties(Range xlRange)
        {
            List<MemoryRow> objs = new List<MemoryRow>();
            int columns = xlRange.Columns.Count;
            int rows = xlRange.Rows.Count;
            int firstColumn = 1;
            int firstRow = 1;
            long lastRow = firstRow + rows - 1;
            MemoryRow memoryHeaders = new MemoryRow()
            {
                CustNo =        "CUSTNO",
                ShipTo =        "SHIP_TO",
                PONO =          "PONO",
                CSR =           "CSR",
                SLSNAME =       "SLSNAME",
                OrderNo =       "ORDERNO",
                ReleaseNo =     "RELEASE_NO",
                OrdDate =       "ORD_DATE",
                PromiseDate =   "PROMISE_DATE",
                ItemNo =        "ITEMNO",
                CustItemNo =    "CUST_ITEMNO",
                CustDescrip =   "CUST_DESCRIP",
                WHSE =          "WHSE",
                OrdQty =        "ORD-QTY",
                OrdAvailQty =   "ORD-AVAIL-QTY",
                HoldTerms =     "HOLD Terms"
            };
            objs.Add(memoryHeaders);

            //get Header row
            var nulvalue = Convert.ToString((xlRange.Cells[4, 5]).Value);
            for (int i = firstRow; i <= lastRow; i++)
            {
                MemoryRow memoryRow = new MemoryRow()
                {
                    CustNo =        Convert.ToString((xlRange.Cells[i, firstColumn]).Value),
                    ShipTo =        Convert.ToString((xlRange.Cells[i, firstColumn + 1]).Value),
                    PONO =          Convert.ToString((xlRange.Cells[i, firstColumn + 2]).Value),
                    CSR =           Convert.ToString((xlRange.Cells[i, firstColumn + 3]).Value),
                    SLSNAME =       Convert.ToString((xlRange.Cells[i, firstColumn + 4]).Value),
                    OrderNo =       Convert.ToString((xlRange.Cells[i, firstColumn + 5]).Value),
                    ReleaseNo =     Convert.ToString((xlRange.Cells[i, firstColumn + 6]).Value),
                    OrdDate =       Convert.ToString((xlRange.Cells[i, firstColumn + 7]).Value),
                    PromiseDate =   Convert.ToString((xlRange.Cells[i, firstColumn + 8]).Value),
                    ItemNo =        Convert.ToString((xlRange.Cells[i, firstColumn + 9]).Value),
                    CustItemNo =    Convert.ToString((xlRange.Cells[i, firstColumn + 10]).Value),
                    CustDescrip =   Convert.ToString((xlRange.Cells[i, firstColumn + 11]).Value),
                    WHSE =          Convert.ToString((xlRange.Cells[i, firstColumn + 12]).Value),
                    OrdQty =        Convert.ToString((xlRange.Cells[i, firstColumn + 13]).Value),
                    OrdAvailQty =   Convert.ToString((xlRange.Cells[i, firstColumn + 14]).Value),
                    HoldTerms =     Convert.ToString((xlRange.Cells[i, firstColumn + 15]).Value)
                };
                //TestOutOfRange = Convert.ToString((xlRange.Cells[222, 57]).Value)
                if (memoryRow.CustNo != "CUSTNO" && memoryRow.CustNo != " Alpha Open Orders Report (CSR)")
                {
                    objs.Add(memoryRow);
                }
                
            }
            //remove empty columns
                //get  headers row.
            return objs;

        }

    }
}
