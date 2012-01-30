using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model.CodaxyRepository;

namespace Tests.CodaxyLib
{
    [TestClass]
    public class TestOpenXMLExcel : TestExcelDataLoader
    {
        [TestInitialize]
        public new void Initialize()
        {
            ReaderRepository = new CodaxyExcelDataReaderRepository();
        }

        [TestMethod]
        [ExpectedException(typeof(ApplicationException), "XlioProvider dont't support the BIFF Format")]
        public void LoadWorkbook_WithBiffFormat_ReturnException()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read))
            {
                ReaderRepository.LoadBinaryWorkbook(stream);
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