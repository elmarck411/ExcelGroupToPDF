﻿using System;
using Microsoft.Office.Interop;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using iText.Kernel.Pdf;
using iText.Kernel.Font;
using iText.IO.Font;
using iText.Layout.Properties;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Geom;
using iText.IO.Font.Constants;
using Path = System.IO.Path;
using System.Globalization;
using System.Diagnostics;
using System.Configuration;

namespace ExcelGroupToPDF
{
    
    class Program
    {
        public static readonly string DEST = "colored_background.pdf";
        public const int NumberOfColumns = 16;
       
        static void Main(string[] args)
        {
            
            Application exApp = null;
            string spreadsheetLocation = null;
            Workbook exWbk = null;
            Worksheet exWks = null;
            Range xlRange = null;
            try
            {
                exApp = new Application();
                
                spreadsheetLocation = Path.Combine(ConfigurationManager.AppSettings["ExcelOriginPath"] + ConfigurationManager.AppSettings["ExcelOriginFileName"]);
                Trace.WriteLine(DateTime.Now.ToString() + " - TESTING CHANGE BEFORE READING EXCEL, spreadsheetLocation: " +  spreadsheetLocation);
                exWbk = exApp.Workbooks.Open(spreadsheetLocation);
                exWks = exWbk.Sheets["Sheet1"];
                xlRange = exWks.UsedRange;
                exWks.EnableAutoFilter = true;
                
                MemoryRow memoryHeaders = new MemoryRow()
                {
                    CustNo = "Cust No",
                    ShipTo = "Ship To",
                    PONO = "PO No.",
                    CSR = "CSR",
                    SLSNAME = "SLS Name",
                    OrderNo = "Order No.",
                    ReleaseNo = "Release No.",
                    OrdDate = "Ord. Date",
                    PromiseDate = "Promise Date",
                    ItemNo = "Item No.",
                    CustItemNo = "Cust. Item No.",
                    CustDescrip = "Cust. Description",
                    WHSE = "WHSE",
                    OrdQty = "Ord. Qty",
                    OrdAvailQty = "Ord. Avail. Qty.",
                    HoldTerms = "Hold Terms"
                };
                
                List<MemoryRow> objs = assignProperties(xlRange);
                var groupedFields = (from o in objs
                                     group o by o.CustNo);
                Trace.WriteLine(DateTime.Now.ToString() + " - Directory path:" + ConfigurationManager.AppSettings["DirPdf"]);
                Directory.CreateDirectory(ConfigurationManager.AppSettings["DirPdf"]);

                Trace.WriteLine(DateTime.Now.ToString() + " - Directory created, entering groups");
                foreach (var custNoGroup in groupedFields)
                {
                    if (custNoGroup.Key != "" && custNoGroup.Key != "Alpha Open Orders Report (CSR)")
                    {
                        // Save into a PDF.
                        #region savepdf
                        Trace.WriteLine(DateTime.Now.ToString() + " - Saving to PDF");
                        PdfDocument pdfDoc = new PdfDocument(new PdfWriter(Path.Combine(ConfigurationManager.AppSettings["DirPdf"],"_ALP_OpenOrderCSR_" + custNoGroup.Key + ".pdf")));
                        Document doc = new Document(pdfDoc, PageSize.LEGAL.Rotate());
                        PdfFont font = PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN);
                        Table table = new Table(UnitValue.CreatePercentArray(new float[] { 13, 7, 5, 9, 4, 3, 4, 4, 4, 13, 18, 2, 5, 6, 3 })).UseAllAvailableWidth();

                        table.AddCell(new Cell(1, 3).Add(new Paragraph("Alpha Open Orders Report")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFontSize(14));
                        table.StartNewRow();
                        table.AddCell(new Cell().Add(new Paragraph("Customer: " + custNoGroup.Key)).SetBorder(iText.Layout.Borders.Border.NO_BORDER));
                        table.StartNewRow();
                        table.StartNewRow();
                        table = CreateTableRow(table, memoryHeaders, true);

                        foreach (var row in custNoGroup)
                        {
                            table = CreateTableRow(table, row);
                        }
                        table.SetFontSize(8);
                        doc.Add(table);
                        doc.Close();

                        #endregion
                        Trace.WriteLine(DateTime.Now.ToString() + " - PDF saved: " + custNoGroup.Key);
                    }
                }
                Trace.WriteLine(DateTime.Now.ToString() + " - App has run correctly");
                
            }
            catch (Exception ex)
            {
                Trace.TraceError(DateTime.Now.ToString() + " - Error: " + ex.Message);
                Trace.TraceError("   - StackTrace: " + ex.StackTrace);

            }
            finally
            {
                if (exApp != null)
                {
                    if(exApp.Workbooks != null)
                        exApp.Workbooks.Close();
                    exApp.Quit();
                }
                Trace.Flush();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="table"></param>
        /// <param name="row"></param>
        /// <param name="isHeader"></param>
        /// <returns></returns>
        public static Table CreateTableRow(Table table, MemoryRow row, bool isHeader = false )
        {
            PdfFont font = PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN);
            if (isHeader)
            {
                font = PdfFontFactory.CreateFont(StandardFonts.TIMES_BOLD);
            }

            table.AddCell(new Cell().Add(new Paragraph(row.ShipTo ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.PONO ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.CSR ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.SLSNAME ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.OrderNo ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.ReleaseNo ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));

            string ordDate = "";
            string promiseDate = "";
            if (isHeader)
            {
                ordDate = row.OrdDate ?? "";
                promiseDate = row.PromiseDate ?? "";
            }
            else
            {
                CultureInfo culture = new CultureInfo("en-US");
                if (!string.IsNullOrEmpty(row.OrdDate))
                {
                    ordDate = Convert.ToDateTime(row.OrdDate, culture).ToShortDateString();
                }

                if (!string.IsNullOrEmpty(row.PromiseDate))
                {
                    promiseDate = Convert.ToDateTime(row.PromiseDate, culture).ToShortDateString();
                }
            }

            table.AddCell(new Cell().Add(new Paragraph(ordDate)).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(promiseDate)).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.ItemNo ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.CustItemNo ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.CustDescrip ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.WHSE ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.OrdQty ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.OrdAvailQty ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));
            table.AddCell(new Cell().Add(new Paragraph(row.HoldTerms ?? "")).SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetFont(font));

            table.StartNewRow();
            return table;
        }


        public static List<MemoryRow>  assignProperties(Range xlRange)
        {
            List<MemoryRow> objs = new List<MemoryRow>();
            int columns = xlRange.Columns.Count;
            int rows = xlRange.Rows.Count;
            int firstColumn = 1;
            int firstRow = 1;
            long lastRow = firstRow + rows - 1;
           

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
