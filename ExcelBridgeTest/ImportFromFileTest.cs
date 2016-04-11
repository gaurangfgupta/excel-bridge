using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelBridge;

namespace ExcelBridgeTest
{
    [TestClass]
    public class ImportFromFileTest
    {
        [TestMethod]
        public void ImportTest()
        {
            ExcelFile xl = new ExcelFile(@"C:\Users\Administrator.GCCHR\Documents\Contact list - Copy.xlsx");
            System.Data.DataSet importedData = xl.Import();
        }
    }
}
