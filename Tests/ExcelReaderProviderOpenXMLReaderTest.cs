using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model.Providers;

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