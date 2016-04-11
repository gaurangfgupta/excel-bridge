using System;
using System.Collections.Generic;
using System.Data;
//Excel
using Microsoft.Office.Interop.Excel;

namespace ExcelBridge
{
    /// <summary>
    /// Import data from Excel file.
    /// </summary>
    public class ExcelFile : IDisposable
    {
        private Application _excelApp;
        private Workbook _book;
        private Sheets _allSheets;
        private DataSet _completeData;

        public ExcelFile(string filePath)
        {
            _excelApp = new Application();
            _book = _excelApp.Workbooks.Open(filePath);
            _allSheets = _book.Sheets;
            //completeData = new DataSet();
        }

        ~ExcelFile()
        {
            Dispose();
        }

        public void Dispose()
        {
            _completeData = null;

            //Release Excel sheets
            if (_allSheets!=null)
            {
                releaseObject(_allSheets);
                _allSheets = null; 
            }

            //Release close Excel book/file
            if (_book != null)
            {
                _book.Close(false);
                releaseObject(_book);
                _book = null;
            }
            //Quit MS Excel
            if (_excelApp != null)
            {
                _excelApp.Quit();
                releaseObject(_excelApp);
                _excelApp = null;
            }

        }

        /// <summary>
        /// Import data from Excel file.
        /// <para>No validations are performed.</para>
        /// </summary>
        /// <returns><see cref="DataSet"/> corresponding to the Excel file, containing one or more <see cref="System.Data.DataTable"/> corresponding to each Excel sheet.
        /// Each row of table corresponds to each row of Excel sheet</returns>
        /// <exception cref="DuplicateNameException">Thrown if Excel contains sheets with duplicate name</exception>
        public DataSet Import()
        {
            _completeData = new DataSet();

            foreach (Worksheet sheet in _allSheets)
            {
                System.Data.DataTable table = new System.Data.DataTable();
                string[] headers;

                //System.Diagnostics.Debug.Write("**************************************************************************");
                //System.Diagnostics.Debug.Write(sheet.Name);
                //System.Diagnostics.Debug.WriteLine("**************************************************************************");

                //Add a table for each sheet to DataSet of all contacts and name it same as the name of the sheet
                table.TableName = sheet.Name;

                //Adding headers to table. Steps:
                //1. Get headers from Excel sheet
                //2. Add the headers obtained from Step 1 to the table
                headers = getHeaders(sheet);
                addHeadersToTable(headers, ref table);
                headers = null;

                //Import data from sheet into table
                addValuesToTable(sheet, ref table);

                try
                {
                    _completeData.Tables.Add(table);
                }
                catch (DuplicateNameException)
                {

                    throw;
                }
                table = null;
                releaseObject(sheet);
            }
            //_book.Close(false);
            //_excelApp.Quit();
            //releaseObject(allSheets);
            //releaseObject(_book);
            //releaseObject(_excelApp);

            return _completeData;
        }

        /// <summary>
        /// Get the headers Excel sheet, assuming that Row 1 of Excel file contains headers, beginning from Column A with no blank columns in between the first and last header
        /// </summary>
        /// <param name="fromSheet">The Excel sheet</param>
        /// <returns>a list of headers</returns>
        private string[] getHeaders(Worksheet fromSheet)
        {
            //Get headers from the sheet
            Range usedRangeInSheet = fromSheet.UsedRange;
            Range row = usedRangeInSheet.Rows[1];
            string[] headers = new string[row.Columns.Count];

            if (row.Row == 1)
            {
                int i = 0;
                foreach (Range col in row.Columns)
                {
                    headers[i] = col.Text;
                    i++;
                    releaseObject(col);
                }
            }
            releaseObject(row);
            releaseObject(usedRangeInSheet);
            return headers;
        }

        /// <summary>
        /// Add columns from the <paramref name="_headers"/> list to the provided <paramref name="_table"/>
        /// </summary>
        /// <param name="_headers">The list of headers to be added to table</param>
        /// <param name="_table">The table in which columns are to be created</param>
        private void addHeadersToTable(string[] _headers, ref System.Data.DataTable _table)
        {
            foreach (string header in _headers)
            {
                _table.Columns.Add(header);
            }
        }

        /// <summary>
        /// Fills the <paramref name="_table"/> with data from the <paramref name="_sheet"/>
        /// </summary>
        /// <param name="_sheet">The Excel sheet to get data from</param>
        /// <param name="_table">The table where the data is to be filled</param>
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
                        //System.Diagnostics.Debug.Write("\t");
                        //System.Diagnostics.Debug.Write(value);
                        //System.Diagnostics.Debug.WriteLine("<<<<<Row End");
                        releaseObject(column);
                    }
                    _table.Rows.Add(newRow);
                    newRow = null;
                }
                releaseObject(row);
            }
            releaseObject(used);
        }

        /// <summary>
        /// Releases resources of the Excel objects like rows, columns, sheet, application etc.
        /// </summary>
        /// <param name="obj">The object to release</param>
        private static void releaseObject(object obj)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
            GC.Collect();
        }

        /// <summary>
        /// Assuming that Row 1 of Excel file contains headers, beginning from Column A with no blank columns in between the first and last header
        /// </summary>
        /// <param name="fromSheetName">The name of the sheet from which to get headers</param>
        /// <returns>Headers from specified sheet</returns>
        public string[] GetHeaders(string fromSheetName)
        {
            _allSheets = _book.Sheets;
            string[] headers = getHeaders(_allSheets.Item[fromSheetName]);
            return headers;
        }

        /// <summary>
        /// Returns 
        /// </summary>
        /// <returns>Headers from all sheets of Excel file in dicionary where Key contains the sheet name and values are array of headers for that Key</returns>
        public Dictionary<string, string[]> GetHeaders()
        {
            Dictionary<string, string[]> sheetHeaders = new Dictionary<string, string[]>();
            _allSheets = _book.Sheets;
            foreach (Worksheet sheet in _allSheets)
            {
                sheetHeaders.Add(sheet.Name, getHeaders(sheet));
            }
            return sheetHeaders;
        }

        public bool HeadersExist(string sheetName, string[] referenceHeaders)
        {
            //Dictionary<string, string> offenders = new Dictionary<string, string>();
            bool exists = true;
            string[] headers = GetHeaders(sheetName);

            foreach (string referenceHeader in referenceHeaders)
            {
                bool headerExists = Array.IndexOf(headers, referenceHeader) > -1;
                if (!headerExists)
                {
                    //offenders.Add(referenceHeader, table.TableName);
                    exists = false;
                }
                if (!exists) break;
            }
            return exists;
        }

        public bool HeadersExist(string[] referenceHeaders)
        {
            //Dictionary<string, string> offenders = new Dictionary<string, string>();
            bool exists = true;

            foreach (string referenceHeader in referenceHeaders)
            {
                foreach (Worksheet sheet in _allSheets)
                {
                    string[] headers = getHeaders(sheet);
                    bool headerExists = Array.IndexOf(headers, referenceHeader) > -1;
                    if (!headerExists)
                    {
                        //offenders.Add(referenceHeader, table.TableName);
                        exists = false;
                    }
                    if (!exists) break;
                }
                if (!exists) break;
            }
            return exists;
        }


    }
}