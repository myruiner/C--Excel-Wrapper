using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model;

namespace Tests
{
    [TestClass]
    public class TestExcelDataLoader_WithXlioProvider : TestExcelDataLoader
    {
        [TestInitialize()]
        public new void Initialize()
        {
            ReaderRepository = new CodaxyExcelDataReaderRepository();
        }

        [TestMethod]
        [ExpectedException(typeof(ApplicationException), "XlioProvider dont't support the BIFF Format")]
        public void LoadBinaryWorkbook_WithBiffFormat_ReturnException()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read))
            {
                ReaderRepository.LoadBinaryWorkbook(stream);
            }
        }
    }
}