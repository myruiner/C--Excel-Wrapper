using ExcelReader.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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