using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model;

namespace Tests
{
    [TestClass]
    public abstract class TestExcelDataLoader
    {
        protected IExcelReaderRepository ReaderRepository;

        [TestCleanup]
        public void CleanUp()
        {
            if (ReaderRepository != null)
            {
                ReaderRepository.Dispose();
            }
        }
    }
}