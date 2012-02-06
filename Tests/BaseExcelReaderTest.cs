using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model;

namespace Tests
{
    [TestClass]
    public abstract class BaseExcelReaderTest
    {
        protected ExcelReader ReaderService;
        protected IExcelReaderProvider Provider;

        [TestInitialize]
        public void SetupExcelReader()
        {
            ReaderService = ExcelReader.ConfigureProvider(Provider);
        }

        [TestMethod]
        public virtual void LoadBinaryWorkbook_WithBinaryFormat_ReturnSuccess()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicBiff.xls", FileMode.Open, FileAccess.Read))
            {
                ReaderService.LoadFile(stream, false);
            }
        }

        [TestMethod]
        public virtual void LoadOpenXMLWorkbook_WithOpenXMLFormat_ReturnSuccess()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicOpenXML.xlsx", FileMode.Open, FileAccess.Read))
            {
                ReaderService.LoadFile(stream, true);
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            if (ReaderService != null)
            {
                ReaderService.Dispose();
            }
        }
    }
}