using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model;

namespace Tests
{
    [TestClass]
    public class TestExcelDataReader
    {
        [TestMethod]
        public void LoadExcelBiff()
        {
            FileStream stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read);
            var repository = new ExcelDataReaderRepository();
            Assert.IsNotNull(repository.LoadBinaryWorkbook(stream));
        }

        [TestMethod]
        public void LoadExcelBiffWrongFormat()
        {
            FileStream stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff_xlsxFormat.xlsx", FileMode.Open, FileAccess.Read);
            var repository = new ExcelDataReaderRepository();
            Assert.IsNotNull(repository.LoadBinaryWorkbook(stream));
        }
    }
}