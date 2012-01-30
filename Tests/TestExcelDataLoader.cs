using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model;

namespace Tests
{
    [TestClass]
    public abstract class TestExcelDataLoader
    {
        protected IExcelReaderRepository ReaderRepository;

        [TestCleanup()]
        public void Initialize()
        {
            if (ReaderRepository != null)
            {
                ReaderRepository.Dispose();
            }
        }

        /// <summary>
        /// Loads the excel biff wrong format.
        /// </summary>
        [TestMethod]
        public void LoadOpenXmlWorkbook__WithOpenXMLFormat_ReturnObject()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff_xlsxFormat.xlsx", FileMode.Open, FileAccess.Read))
            {
                Assert.IsNotNull(ReaderRepository.LoadOpenXmlWorkbook(stream));
            }
        }

        /// <summary>
        /// Loads the excel biff wrong format.
        /// </summary>
        [TestMethod]
        public void LoadOpenXmlWorkbook__WithBiffXMLFormat_ReturnNULL()
        {
            using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read))
            {
                Assert.IsNull(ReaderRepository.LoadOpenXmlWorkbook(stream));
            }
        }

        //[TestMethod]
        //public void GetDataSet_WithBiffFormat_ReturnDataSet()
        //{
        //    using (var stream = File.Open(@"..\..\..\Tests\ExcelTemplates\BasicExcelBiff.xls", FileMode.Open, FileAccess.Read))
        //    {
        //        using (var repository = new ExcelDataReaderRepository())
        //        {
        //            var res = repository.LoadBinaryWorkbook(stream).GetValue(2, 2);
        //            var res2 = repository.LoadBinaryWorkbook(stream).GetValue(2, 3);
        //        }
        //    }
        //}
    }
}