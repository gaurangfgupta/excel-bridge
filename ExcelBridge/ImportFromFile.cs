using System;
using System.Collections.Generic;
using System.Data;
//Excel
using Microsoft.Office.Interop.Excel;

namespace ExcelUtilities.WorkerClasses
{
    public class ImportFromFile
    {
        private Application _excelApp;
        private Workbook _book;
        private Sheets allSheets;
        private DataSet completeData;

        public ImportFromFile(string filePath)
        {
            _excelApp = new Application();
            _book = _excelApp.Workbooks.Open(filePath);
            //completeData = new DataSet();
        }

        ~ImportFromFile()
        {
            _excelApp = null; ;
            _book = null;
            allSheets = null;
            completeData = null;

        }

        public DataSet Import()
        {
            completeData = new DataSet();
            allSheets = _book.Sheets;
            foreach (Worksheet sheet in allSheets)
            {
                System.Data.DataTable table = new System.Data.DataTable();
                List<string> headers;

                //System.Diagnostics.Debug.Write("**************************************************************************");
                //System.Diagnostics.Debug.Write(sheet.Name);
                //System.Diagnostics.Debug.WriteLine("**************************************************************************");

                //Add a table for each sheet to DataSet of all contacts and name it same as the name of the sheet
                table.TableName = sheet.Name;

                //Adding headers to table
                headers = getHeaders(sheet);
                addHeadersToTable(headers, ref table);
                headers = null;

                //Import data from sheet into table
                addValuesToTable(sheet, ref table);
                completeData.Tables.Add(table);
                table = null;
                releaseObject(sheet);
            }
            _book.Close(false);
            _excelApp.Quit();
            releaseObject(_book);
            releaseObject(_excelApp);
            
            return completeData;
        }

        private List<string> getHeaders(Worksheet fromSheet)
        {
            List<string> headers = new List<string>();
            //Get headers from the sheet
            Range usedRangeInSheet = fromSheet.UsedRange;
            foreach (Range row in usedRangeInSheet.Rows)
            {
                if (row.Row == 1)
                {
                    foreach (Range col in row.Columns)
                    {
                        headers.Add(col.Text);
                        releaseObject(col);
                    }
                }
                releaseObject(row);
            }
            releaseObject(usedRangeInSheet);
            return headers;
        }

        private void addHeadersToTable(List<string> _headers, ref System.Data.DataTable _table)
        {
            foreach (string header in _headers)
            {
                _table.Columns.Add(header);
            }
        }

        private void addValuesToTable(Worksheet _sheet, ref System.Data.DataTable _table)
        {
            Range used = _sheet.UsedRange;
            used.ClearFormats();
            foreach (Range row in used.Rows)
            {
                if (row.Row > 1)
                {
                    int i = 0;
                    DataRow newRow = _table.NewRow();

                    foreach (Range column in row.Columns)
                    {
                        string value;
                        //value = (column.Text is DBNull) ? "-----" : column.Text;
                        value = column.Text ?? null;
                        newRow[i++] = value;
                        System.Diagnostics.Debug.Write("\t");
                        System.Diagnostics.Debug.Write(value);
                        System.Diagnostics.Debug.WriteLine("<<<<<Row End");
                        releaseObject(column);
                    }
                    _table.Rows.Add(newRow);
                    newRow = null;
                }
                releaseObject(row);
            }
            releaseObject(used);
        }

        private static void releaseObject(object obj)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
            GC.Collect();
        }
    }
}