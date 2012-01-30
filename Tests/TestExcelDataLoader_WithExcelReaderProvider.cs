using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model;

namespace Tests
{
    [TestClass]
    public class TestExcelDataLoader_WithExcelReaderProvider : TestExcelDataLoader
    {
        [TestInitialize()]
        public new void Initialize()
        {
            ReaderRepository = new ExcelDataReaderRepository();
        }

        [TestMethod]
        public void LoadBinaryWorkbook_WithBiffFormat_ReturnObject()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read))
            {
                Assert.IsNotNull(ReaderRepository.LoadBinaryWorkbook(stream));
            }
        }

        [TestMethod]
        public void LoadBinaryWorkbook_WithOpenXMLFormat_ReturnNULL()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff_xlsxFormat.xlsx", FileMode.Open, FileAccess.Read))
            {
                Assert.IsNull(ReaderRepository.LoadBinaryWorkbook(stream));
            }
        }
    }
}