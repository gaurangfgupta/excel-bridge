using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ExcelBridge.Helper
{
    public static class DataSetHelper
    {
        public static Dictionary<string, bool> CheckColumnsConsistency(DataSet dataset, bool ignoreColumnOrder = true)
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

        private static List<string> getColumnsOfDataTable(DataTable sourceTable, bool sort = false)
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

        public static Dictionary<string, int> RemoveBlankRows(ref DataSet dataSet)
        {
            Dictionary<string, int> deletionSummary;
            deletionSummary = new Dictionary<string, int>();

            foreach (DataTable table in dataSet.Tables)
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

    }
}
