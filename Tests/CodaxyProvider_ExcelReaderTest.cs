using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model.Providers;

namespace Tests
{
    [TestClass]
    public class CodaxyProvider_ExcelReaderTest : BaseExcelReaderTest
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CodaxyProvider_ExcelReaderTest"/> class.
        /// </summary>
        public CodaxyProvider_ExcelReaderTest()
        {
            Provider = new CodaxyExcelProvider();
        }

        [ExpectedException(typeof(ApplicationException), "Binary Format not supported by CodaxyExcelDataReader")]
        [TestMethod]
        public override void LoadBinaryWorkbook_WithBinaryFormat_ReturnSuccess()
        {
            base.LoadBinaryWorkbook_WithBinaryFormat_ReturnSuccess();
        }
    }
}