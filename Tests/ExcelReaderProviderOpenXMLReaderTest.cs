using ExcelReader.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests
{
    [TestClass]
    public class ExcelReaderProviderOpenXMLReaderTest : BaseExcelReaderTest
    {
        public ExcelReaderProviderOpenXMLReaderTest()
        {
            Provider = new ExcelDataReaderProvider();
            IsOpenXML = true;
        }
    }
}