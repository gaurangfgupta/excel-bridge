using System;
using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelBridge.Helper;
using System.Collections.Generic;

namespace ExcelBridgeTest
{
    [TestClass]
    public class DatatableValidatorTest
    {
        [TestMethod]
        public void CheckColumnsConsistency()
        {
            DataSet _dataset;
            _dataset = new DataSet();
            Dictionary<string, bool> resultSummary;

            #region Create reference table to compare other tables with
            //Create reference table and add to dataset
            DataTable referenceTable = new DataTable("Reference Table");
            for (int i = 1; i < 6; i++)
            {
                referenceTable.Columns.Add(string.Format("Column{0}", i));
            }
            _dataset.Tables.Add(referenceTable);
            //Check consistency
            resultSummary = DataSetHelper.CheckColumnsConsistency(_dataset);
            //Assert expected results
            Assert.AreEqual(resultSummary.Count, 0);
            //Release objects
            #endregion

            #region Check equal table
            //Create target table and add to dataset
            DataTable equalTable = new DataTable("Equivalent Table");
            for (int i = 1; i < 6; i++)
            {
                equalTable.Columns.Add(string.Format("Column{0}", i));
            }
            _dataset.Tables.Add(equalTable);
            //Check consistency ignoring column order
            resultSummary = DataSetHelper.CheckColumnsConsistency(_dataset);
            //Assert expected results
            foreach (bool columnConsistencyResult in resultSummary.Values)
            {
                Assert.IsTrue(columnConsistencyResult);
            }
            //Check consistency considering column order
            resultSummary = DataSetHelper.CheckColumnsConsistency(_dataset,false);
            //Assert expected results
            foreach (bool columnConsistencyResult in resultSummary.Values)
            {
                Assert.IsTrue(columnConsistencyResult);
            }
            //Release objects
            equalTable = null;
            _dataset.Tables.RemoveAt(1);
            #endregion

            #region Check table with same columns but different order
            //Create target table and add to dataset
            DataTable differentOrderTable = new DataTable("Different order table");
            for (int i = 5; i > 0; i--)
            {
                differentOrderTable.Columns.Add(string.Format("Column{0}", i));
            }
            _dataset.Tables.Add(differentOrderTable);
            //Check consistency ignoring column order
            resultSummary = DataSetHelper.CheckColumnsConsistency(_dataset);
            //Assert expected results
            Assert.IsTrue(resultSummary[differentOrderTable.TableName]);
            //Check consistency considering column order
            resultSummary = DataSetHelper.CheckColumnsConsistency(_dataset,false);
            //Assert expected results
            Assert.IsFalse(resultSummary[differentOrderTable.TableName]);
            //Release objects
            differentOrderTable = null;
            _dataset.Tables.RemoveAt(1);
            #endregion

            #region Check table with different column count
            //Create target table and add to dataset
            DataTable differentCountTable = new DataTable("Different column count table");
            for (int i = 1; i < 7; i++)
            {
                differentCountTable.Columns.Add(string.Format("Column{0}", i));
            }
            _dataset.Tables.Add(differentCountTable);
            //Check consistency
            resultSummary = DataSetHelper.CheckColumnsConsistency(_dataset);
            //Assert expected results
            Assert.IsFalse(resultSummary[differentCountTable.TableName]);
            //Release objects
            differentCountTable = null;
            _dataset.Tables.RemoveAt(1);
            #endregion

            #region Check table with different columns
            //Create target table and add to dataset
            DataTable differentColumnsTable = new DataTable("Different column count table");
            for (int i = 1; i < 5; i++)
            {
                differentColumnsTable.Columns.Add(string.Format("Column{0}", i));
            }
            differentColumnsTable.Columns.Add(string.Format("Column{0}", 7));
            _dataset.Tables.Add(differentColumnsTable);
            //Check consistency ignoring column order
            resultSummary = DataSetHelper.CheckColumnsConsistency(_dataset);
            //Assert expected results
            Assert.IsFalse(resultSummary[differentColumnsTable.TableName]);
            //Check consistency considering column order
            resultSummary = DataSetHelper.CheckColumnsConsistency(_dataset,false);
            //Assert expected results
            Assert.IsFalse(resultSummary[differentColumnsTable.TableName]);
            //Release objects
            differentColumnsTable = null;
            _dataset.Tables.RemoveAt(1);
            #endregion
        }

        [TestMethod]
        public void GetRecordCount()
        {
            DataSet ds = new DataSet();

            DataTable table1 = ds.Tables.Add();
            table1.TableName = "Table 1";
            table1.Columns.Add("Values");
            for (int i = 1; i <= 10; i++)
            {
                DataRow row = table1.NewRow();
                row[0] = string.Format("Value # {0}", i);
                table1.Rows.Add(row);
            }

            DataTable table2 = ds.Tables.Add();
            table2.TableName = "Table 2";
            table2.Columns.Add("Values");
            for (int i = 1; i <= 100; i++)
            {
                DataRow row = table2.NewRow();
                row[0] = string.Format("Value # {0}", i);
                table2.Rows.Add(row);
            }

             Dictionary<string,int> recordCount= DataSetHelper.GetRecordsCount(ref ds);
            Assert.AreEqual(10, recordCount[table1.TableName]);
            Assert.AreEqual(100, recordCount[table2.TableName]);
        }
    }

}
