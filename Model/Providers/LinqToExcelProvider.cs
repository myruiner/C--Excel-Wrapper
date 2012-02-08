// -----------------------------------------------------------------------
// <copyright file="LinqToExcelProvider.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System.Data;
using System.IO;
using System.Linq;
using LinqToExcel;
using LinqToExcel.Domain;
using Model.Extensions;
using Model.Frame;

namespace Model.Providers
{
    using System;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class LinqToExcelProvider : IExcelReaderProvider
    {
        private ExcelQueryFactory _excelQueryFactory;
        private TemporaryFileInfo _temporaryFileInfo;
        private int _currentActiveWorksheet;

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            _temporaryFileInfo.Dispose();
        }

        /// <summary>
        /// Gets the content of cell identified by ID
        /// </summary>
        /// <param name="column">The column.</param>
        /// <param name="row">The row.</param>
        /// <returns></returns>
        public object GetValueFromCellByID(int column, int row)
        {
            var query = from c in _excelQueryFactory.WorksheetNoHeader(_currentActiveWorksheet)
                        select c[column];
            if (query.Count() == 0 || query.Count() < row)
                return DBNull.Value;
            return query.ToList().ElementAt(row).Value;
        }

        /// <summary>
        /// Gets the content of cell identified by his ID
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public object GetValueFromCellByAddress(string address)
        {
            var query = from c in _excelQueryFactory.WorksheetRangeNoHeader(address, address, _currentActiveWorksheet)
                        select c;

            return null;
        }

        public object GetValueFromCellByName(string namedCell)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromRangeByID(int columnStart, int rowStart, int colunmnEnd, int rowEnd)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromRangeByAddress(string address)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromRangeByName(string namedRange)
        {
            throw new NotImplementedException();
        }

        public DataTable GetWorksheetContent()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Loads from file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        private IExcelReaderProvider LoadFromFile(FileStream stream)
        {
            _temporaryFileInfo = stream.CopyStreamToTempFileInfo();
            _excelQueryFactory = new ExcelQueryFactory(_temporaryFileInfo.GetFileFinfo().FullName);
            return this;
        }

        /// <summary>
        /// Loads Content from a Binary file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public IExcelReaderProvider LoadFromBinaryFile(FileStream stream)
        {
            _excelQueryFactory.DatabaseEngine = DatabaseEngine.Jet;
            return LoadFromFile(stream);
        }

        /// <summary>
        /// Loads content from an OpenXML file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public IExcelReaderProvider LoadFromOpenXMLFile(FileStream stream)
        {
            _excelQueryFactory.DatabaseEngine = DatabaseEngine.Ace;
            return LoadFromFile(stream);
        }

        public IExcelReaderProvider SetCurrentWorksheet(string name)
        {
            _currentActiveWorksheet = 0;
            return this;
        }

        public IExcelReaderProvider SetCurrentWorksheet(int index)
        {
            _currentActiveWorksheet = index;
            return this;
        }

        public DataSet ToDataSet()
        {
            throw new NotImplementedException();
        }

        public IExcelReaderProvider WithFirstRowAsColumnName(bool isFirstRowAsColumnName)
        {
            throw new NotImplementedException();
        }
    }
}