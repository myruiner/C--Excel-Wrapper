// -----------------------------------------------------------------------
// <copyright file="LinqToExcelProvider.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ExcelReader.Extensions;
using ExcelReader.Frame;
using LinqToExcel;
using Row = LinqToExcel.Row;

namespace ExcelReader.Providers
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class LinqToExcelProvider : IExcelReaderProvider
    {
        /// <summary>
        /// Return single value from Query
        /// </summary>
        private readonly Func<IQueryable<Cell>, int, object> _returnSingleValueFromQuery = (query, i) =>
                                                                                               {
                                                                                                   if (query.Count() == 0 || query.Count() < i)
                                                                                                       return DBNull.Value;
                                                                                                   return
                                                                                                       query.ToList().ElementAt(i).Value;
                                                                                               };

        private int _currentActiveWorksheetID;
        private string _currentActiveWorksheetName;
        private ExcelQueryFactory _excelQueryFactory;
        private TemporaryFileInfo _temporaryFileInfo;

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_temporaryFileInfo != null)
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
            var query = from c in (IQueryable<IList<Cell>>)_excelQueryFactory.WorksheetNoHeader(GetCurrentWorksheet())
                        select c[column];
            return _returnSingleValueFromQuery(query, row);
        }

        /// <summary>
        /// Gets the current worksheet.
        /// </summary>
        /// <returns></returns>
        private dynamic GetCurrentWorksheet()
        {
            return string.IsNullOrEmpty(_currentActiveWorksheetName)
                       ? (dynamic)_currentActiveWorksheetID
                       : _currentActiveWorksheetName;
        }

        /// <summary>
        /// Gets the content of cell identified by his ID
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public object GetValueFromCellByAddress(string address)
        {
            var query =
                from c in (IQueryable<IList<Cell>>)_excelQueryFactory.WorksheetRangeNoHeader(address, address, GetCurrentWorksheet())
                select c;
            return query.Count() > 0 ? query.ToList().ElementAt(0)[0].Value : DBNull.Value;
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

        public IEnumerable GetWorksheetContent()
        {
            IQueryable<Row> query = from c in _excelQueryFactory.Worksheet(_currentActiveWorksheetID)
                                    select c;
            return query.ToList();
        }

        /// <summary>
        /// Loads Content from a Binary file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public IExcelReaderProvider LoadFromBinaryFile(FileStream stream)
        {
            _temporaryFileInfo = stream.CopyStreamToTempAndReturnFileInfo("xlsx");
            _excelQueryFactory = new ExcelQueryFactory(_temporaryFileInfo.GetFileFinfo().FullName);
            return this;
        }

        /// <summary>
        /// Loads content from an OpenXML file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public IExcelReaderProvider LoadFromOpenXMLFile(FileStream stream)
        {
            _temporaryFileInfo = stream.CopyStreamToTempAndReturnFileInfo("xlsx");
            _excelQueryFactory = new ExcelQueryFactory(_temporaryFileInfo.GetFileFinfo().FullName);
            return this;
        }

        public IExcelReaderProvider SetCurrentWorksheet(string name)
        {
            _currentActiveWorksheetName = name;
            return this;
        }

        public IExcelReaderProvider SetCurrentWorksheet(int index)
        {
            _currentActiveWorksheetID = index;
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