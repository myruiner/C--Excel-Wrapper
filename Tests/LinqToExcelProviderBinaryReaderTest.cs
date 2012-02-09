using ExcelReader.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests
{
    [TestClass]
    public class LinqToExcelProviderBinaryReaderTest : BaseExcelReaderTest
    {
        public LinqToExcelProviderBinaryReaderTest()
        {
            Provider = new LinqToExcelProvider();
            IsOpenXML = false;
        }

        [TestMethod]
        public override void LoadWorkbookWithWrongFormat_ReturnException()
        {
        }
    }
}