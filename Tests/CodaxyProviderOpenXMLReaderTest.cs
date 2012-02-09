using System;
using ExcelReader.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests
{
    [TestClass]
    public class CodaxyProviderOpenXMLReaderTest : BaseExcelReaderTest
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CodaxyProviderOpenXMLReaderTest"/> class.
        /// </summary>
        public CodaxyProviderOpenXMLReaderTest()
        {
            Provider = new CodaxyExcelProvider();
            IsOpenXML = true;
        }

        [ExpectedException(typeof(ApplicationException), "Binary Format not supported by CodaxyExcelDataReader")]
        [TestMethod]
        public void LoadBinaryWorkbook_ReturnException()
        {
            IsOpenXML = false;
            base.LoadWorkbook_ReturnSuccess();
            IsOpenXML = true;
        }
    }
}