using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model.Providers;

namespace Tests
{
    [TestClass]
    public class ExcelReaderProvider_ExcelReaderTest : BaseExcelReaderTest
    {
        public ExcelReaderProvider_ExcelReaderTest()
        {
            Provider = new ExcelDataReaderProvider();
        }
    }
}