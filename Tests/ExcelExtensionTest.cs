using System;
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

            Throws<ArgumentException>(() => "1a".ToExcelCell());
            Throws<ArgumentException>(() => "a1a".ToExcelCell());
            Throws<ArgumentException>(() => "1a1".ToExcelCell());
        }

        public static void Throws<TException>(Action blockToExecute) where TException : Exception
        {
            try
            {
                blockToExecute();
            }
            catch (Exception ex)
            {
                Assert.IsTrue(ex.GetType() == typeof(TException), "Expected exception of type " + typeof(TException) + " but type of " + ex.GetType() + " was thrown instead.");
                return;
            }

            Assert.Fail("Expected exception of type " + typeof(TException) + " but no exception was thrown.");
        }
    }
}