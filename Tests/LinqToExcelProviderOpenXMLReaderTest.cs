using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model.Providers;

namespace Tests
{
    [TestClass]
    public class LinqToExcelProviderOpenXMLReaderTest : BaseExcelReaderTest
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="LinqToExcelProviderOpenXMLReaderTest"/> class.
        /// </summary>
        public LinqToExcelProviderOpenXMLReaderTest()
        {
            Provider = new LinqToExcelProvider();
            IsOpenXML = true;
        }
    }
}