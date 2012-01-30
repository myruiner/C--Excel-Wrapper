using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model.ExcelDataRepository;

namespace Tests.ExcelReaderLib
{
    [TestClass]
    public class TestOpenBinaryExcel : TestExcelDataLoader
    {
        [TestInitialize]
        public void Initialize()
        {
            ReaderRepository = new ExcelDataReaderRepository();
        }

        [TestMethod]
        public void LoadWorkbook_WithBiffFormat_ReturnSuccess()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read))
            {
                Assert.IsNotNull(ReaderRepository.LoadBinaryWorkbook(stream));
            }
        }

        [ExpectedException(typeof(ApplicationException), "Wrong FileFormat")]
        [TestMethod]
        public void LoadWorkbook_WithOpenXMLFormat_ReturnException()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff_xlsxFormat.xlsx", FileMode.Open, FileAccess.Read))
            {
                ReaderRepository.LoadBinaryWorkbook(stream);
            }
        }
    }
}