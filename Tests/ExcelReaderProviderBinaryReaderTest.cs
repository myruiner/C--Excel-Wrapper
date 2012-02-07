using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model.Providers;

namespace Tests
{
    [TestClass]
    public class ExcelReaderProviderBinaryReaderTest : BaseExcelReaderTest
    {
        public ExcelReaderProviderBinaryReaderTest()
        {
            Provider = new ExcelDataReaderProvider();
            IsOpenXML = false;
        }
    }
}