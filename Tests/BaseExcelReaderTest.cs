using System;
using System.Configuration;
using System.Globalization;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Model;

namespace Tests
{
    [TestClass]
    public abstract class BaseExcelReaderTest
    {
        protected ExcelReader ReaderService;
        protected IExcelReaderProvider Provider;
        protected Func<ExcelReader, FileStream, IExcelReaderProvider> FileLoader;
        protected bool IsOpenXML;

        [TestInitialize]
        public void SetupExcelReader()
        {
            ReaderService = ExcelReader.ConfigureProvider(Provider);
            FileLoader = (reader, stream) => reader.LoadFile(stream, IsOpenXML);
        }

        [TestMethod]
        public virtual void LoadWorksheetByID_ReturnSuccess()
        {
            using (var stream = File.Open(GetExcelPath("Basic"), FileMode.Open, FileAccess.Read))
            {
                var loadedDT = FileLoader.Invoke(ReaderService, stream).SetCurrentWorksheet(1).GetWorksheetContent();
                Assert.IsNotNull(loadedDT);
            }
        }

        [TestMethod]
        public virtual void LoadWorksheetByName_ReturnSuccess()
        {
            using (var stream = File.Open(GetExcelPath("Basic"), FileMode.Open, FileAccess.Read))
            {
                var loadedDT = FileLoader.Invoke(ReaderService, stream).SetCurrentWorksheet("WS2").GetWorksheetContent();
                Assert.IsNotNull(loadedDT);
            }
        }

        [TestMethod]
        public virtual void LoadWorkbook_ReturnSuccess()
        {
            using (var stream = File.Open(GetExcelPath("Basic"), FileMode.Open, FileAccess.Read))
            {
                Assert.IsNotNull(FileLoader.Invoke(ReaderService, stream));
            }
        }

        [ExpectedException(typeof(ApplicationException), "Wrong Format or Format not supported ba the provider")]
        [TestMethod]
        public virtual void LoadWorkbookWithWrongFormat_ReturnException()
        {
            using (var stream = File.Open(GetExcelPath("Basic"), FileMode.Open, FileAccess.Read))
            {
                FileLoader = (reader, loadedStream) => reader.LoadFile(loadedStream, !IsOpenXML);
                FileLoader.Invoke(ReaderService, stream);
            }
        }

        [TestMethod]
        public virtual void ParseBasicContent_ReturnSucccess()
        {
            using (var stream = File.Open(GetExcelPath("Basic"), FileMode.Open, FileAccess.Read))
            {
                var loadedProvider = FileLoader.Invoke(ReaderService, stream);
                Assert.AreEqual(1, int.Parse(loadedProvider.SetCurrentWorksheet(0).GetValueFromCellByID(0, 0).ToString()));
                Assert.AreEqual(DBNull.Value, loadedProvider.SetCurrentWorksheet(0).GetValueFromCellByID(1, 1));
                Assert.AreEqual(new DateTime(2011, 1, 13), DateTime.ParseExact(loadedProvider.SetCurrentWorksheet(0).GetValueFromCellByID(2, 2).ToString(), "d.M.yyyy", CultureInfo.InvariantCulture));
                Assert.AreEqual("NameCell", loadedProvider.SetCurrentWorksheet(0).GetValueFromCellByID(2, 3).ToString());
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            if (ReaderService != null)
            {
                ReaderService.Dispose();
            }
        }

        /// <summary>
        /// Gets the excel.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        protected string GetExcelPath(string key)
        {
            var pathFile = ConfigurationManager.AppSettings["basePath"];
            var excelFile = ConfigurationManager.AppSettings[key];

            if (IsOpenXML)
                excelFile = excelFile + "x";
            return Path.Combine(pathFile, excelFile);
        }
    }
}