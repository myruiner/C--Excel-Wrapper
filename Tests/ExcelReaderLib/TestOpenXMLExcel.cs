using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model.ExcelDataRepository;

namespace Tests.ExcelReaderLib
{
    [TestClass]
    public class TestOpenXMLExcel : TestExcelDataLoader
    {
        [TestInitialize]
        public new void Initialize()
        {
            ReaderRepository = new ExcelDataReaderRepository();
        }

        [ExpectedException(typeof(ApplicationException), "Wrong FileFormat")]
        [TestMethod]
        public void LoadWorkbook_WithBiffFormat_ReturnException()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read))
            {
                ReaderRepository.LoadOpenXmlWorkbook(stream);
            }
        }

        [TestMethod]
        public void LoadWorkbook_WithOpenXMLFormat_ReturnSuccess()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff_xlsxFormat.xlsx", FileMode.Open, FileAccess.Read))
            {
                Assert.IsNotNull(ReaderRepository.LoadOpenXmlWorkbook(stream));
            }
        }
    }
}