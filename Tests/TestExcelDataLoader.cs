using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model;

namespace Tests
{
    [TestClass]
    public class TestExcelDataLoader
    {
        [TestMethod]
        public void LoadBinaryWorkbook_WithBiffFormat_ReturnObject()
        {
            FileStream stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read);
            var repository = new ExcelDataReaderRepository();
            Assert.IsNotNull(repository.LoadBinaryWorkbook(stream));
        }

        /// <summary>
        /// Loads the excel biff wrong format.
        /// </summary>
        [TestMethod]
        public void LoadBinaryWorkbook_WithOpenXMLFormat_ReturnNULL()
        {
            FileStream stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff_xlsxFormat.xlsx", FileMode.Open, FileAccess.Read);
            var repository = new ExcelDataReaderRepository();
            Assert.IsNull(repository.LoadBinaryWorkbook(stream));
        }

        /// <summary>
        /// Loads the excel biff wrong format.
        /// </summary>
        [TestMethod]
        public void LoadWorkbook__WithOpenXMLFormat_ReturnObject()
        {
            FileStream stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff_xlsxFormat.xlsx", FileMode.Open, FileAccess.Read);
            var repository = new ExcelDataReaderRepository();
            Assert.IsNotNull(repository.LoadWorkbook(stream));
        }

        /// <summary>
        /// Loads the excel biff wrong format.
        /// </summary>
        [TestMethod]
        public void LoadWorkbook__WithOpenXMLFormat_ReturnNULL()
        {
            FileStream stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read);
            var repository = new ExcelDataReaderRepository();
            Assert.IsNull(repository.LoadWorkbook(stream));
        }
    }
}