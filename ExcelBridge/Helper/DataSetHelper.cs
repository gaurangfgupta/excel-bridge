using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ExcelBridge.Helper
{
    public class DataSetHelper
    {
        private DataSet dataset;

        public DataSetHelper(ref DataSet _dataset)
        {
            dataset = _dataset;
        }

        public  Dictionary<string, bool> CheckColumnsConsistency(bool ignoreColumnOrder = true)
        {
            Dictionary<string, bool> validationSummary = new Dictionary<string, bool>();
            if (dataset.Tables.Count > 1)
            {
                #region Prepare reference list of columns from first table of dataset. Used later to compare columns of other tables.
                List<string> referenceColumns = getColumnsOfDataTable(dataset.Tables[0], ignoreColumnOrder);
                #endregion

                foreach (DataTable table in dataset.Tables)
                {
                    bool consistencySwitch = false;
                    List<string> columnsOfTargetTable = getColumnsOfDataTable(table, ignoreColumnOrder);

                    if (referenceColumns.Count == columnsOfTargetTable.Count)
                    {
                        consistencySwitch = referenceColumns.SequenceEqual(columnsOfTargetTable);
                    }

                    validationSummary.Add(table.TableName, consistencySwitch);
                }
            }
            return validationSummary;
        }

        private  List<string> getColumnsOfDataTable(DataTable sourceTable, bool sort = false)
        {
            List<string> columns = new List<string>();
            foreach (DataColumn col in sourceTable.Columns)
            {
                columns.Add(col.ColumnName);
            }
            if (sort)
            {
                columns.Sort();
            }
            return columns;
        }

        public  Dictionary<string, int> RemoveBlankRows()
        {
            Dictionary<string, int> deletionSummary;
            deletionSummary = new Dictionary<string, int>();

            foreach (DataTable table in dataset.Tables)
            {
                int deletedRowsCount = 0;
                List<DataRow> rowsToDelete = new List<DataRow>();

                foreach (DataRow row in table.Rows)
                {
                    //Count blank columns in row
                    int blankColumnCount = 0;
                    foreach (DataColumn col in table.Columns)
                    {
                        if (string.IsNullOrWhiteSpace(row[col.ColumnName].ToString()))
                        {
                            blankColumnCount++;
                        }
                    }
                    //If all columns blank, collect row in deletion list
                    if (blankColumnCount == table.Columns.Count)
                    {
                        rowsToDelete.Add(row);
                        deletedRowsCount++;
                    }
                }

                //Delete collected rows
                foreach (DataRow rowToDelete in rowsToDelete)
                {
                    table.Rows.Remove(rowToDelete);
                }

                deletionSummary.Add(table.TableName, deletedRowsCount);
            }
            return deletionSummary;
        }

        public  Dictionary<string, int> GetRecordsCount()
        {
            Dictionary<string, int> recordCountList = new Dictionary<string, int>();
            foreach (DataTable table in dataset.Tables)
            {
                recordCountList.Add(table.TableName, table.Rows.Count);
            }
            return recordCountList;
        }

        public  bool ColumnsExist(List<string> columns)
        {
            Dictionary<string, string> offenders = new Dictionary<string, string>();
            bool exists = true;

            foreach (string column in columns)
            {
                foreach (DataTable table in dataset.Tables)
                {
                    if (!table.Columns.Contains(column))
                    {
                        offenders.Add(column,table.TableName);
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
