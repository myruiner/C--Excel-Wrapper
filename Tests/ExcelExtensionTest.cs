using ExcelReader.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests
{
    [TestClass]
    public class ExcelExtensionTest
    {
        [TestMethod]
        public void TestToExcelCellExtension()
        {
            Assert.AreEqual(new ExcelCell { Column = 0, Row = 1 }, "A2".ToExcelCell());
            Assert.AreEqual(new ExcelCell { Column = 26, Row = 0 }, "AA1".ToExcelCell());
            Assert.AreEqual(new ExcelCell { Column = 0, Row = 0 }, "a1".ToExcelCell());
        }
    }
}