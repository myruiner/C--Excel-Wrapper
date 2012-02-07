using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model.Providers;

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
    }
}