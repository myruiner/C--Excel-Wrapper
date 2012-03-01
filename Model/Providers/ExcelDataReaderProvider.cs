// -----------------------------------------------------------------------
// <copyright file="ExcelDataReaderRepository.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Collections;
using System.Data;
using System.IO;
using Excel;
using ExcelReader.Extensions;

namespace ExcelReader.Providers
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class ExcelDataReaderProvider : IExcelReaderProvider
    {
        private int _currentActiveWorksheet;
        private DataSet _internalDataSet;
        private bool _isFirstRowAsColumnNames;
        private IExcelDataReader _reader;

        /// <summary>
        /// Loads the internal data set.
        /// </summary>
        private void LoadInternalDataSet()
        {
            _reader.IsFirstRowAsColumnNames = _isFirstRowAsColumnNames;
            _internalDataSet = _reader.AsDataSet();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_reader == null) return;
            _reader.Close();
            _reader.Dispose();
        }

        /// <summary>
        /// Gets the content of cell identified by ID
        /// </summary>
        /// <param name="column">The column.</param>
        /// <param name="row">The row.</param>
        /// <returns></returns>
        public object GetValueFromCellByID(int column, int row)
        {
            if (_internalDataSet == null)
                throw new ApplicationException("Internal DataSet is null");

            try
            {
                return _internalDataSet.Tables[_currentActiveWorksheet].Rows[row][column];
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the value from cell by address.
        /// </summary>
        /// <param name="cellAddress">The cell address.</param>
        /// <returns></returns>
        public object GetValueFromCellByAddress(string cellAddress)
        {
            var cell = cellAddress.ToExcelCell();
            return GetValueFromCellByID(cell.Column, cell.Row);
        }

        /// <summary>
        /// Gets the name of the value from cell by.
        /// </summary>
        /// <param name="columnStart">The column start.</param>
        /// <returns></returns>
        public object GetValueFromCellByName(string columnStart)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the content of a range identified by his ID
        /// </summary>
        /// <param name="columnStart">The column start.</param>
        /// <param name="rowStart">The row start.</param>
        /// <param name="colunmnEnd">The colunmn end.</param>
        /// <param name="rowEnd">The row end.</param>
        /// <returns></returns>
        public DataTable GetValueFromRangeByID(int columnStart, int rowStart, int colunmnEnd, int rowEnd)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromRangeByAddress(string address)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the content of a range identified by his name
        /// </summary>
        /// <param name="namedRange">The named range.</param>
        /// <returns></returns>
        public DataTable GetValueFromRangeByName(string namedRange)
        {
            throw new ApplicationException("This provider does not support named ranges");
        }

        /// <summary>
        /// Return the current speicified Worksheet as a DataTable
        /// </summary>
        /// <returns></returns>
        public IEnumerable GetWorksheetContent()
        {
            if (_internalDataSet == null)
                throw new ApplicationException("Internal DataSet is null");

            return _internalDataSet.Tables[_currentActiveWorksheet].AsEnumerable();
        }

        /// <summary>
        /// Loads Content from a Binary file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public IExcelReaderProvider LoadFromBinaryFile(FileStream stream)
        {
            _reader = ExcelReaderFactory.CreateBinaryReader(stream);
            if (!_reader.IsValid)
            {
                throw new ApplicationException("BIFF Format not recognized");
            }
            LoadInternalDataSet();
            return this;
        }

        /// <summary>
        /// Sets the current worksheet.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <returns></returns>
        public IExcelReaderProvider SetCurrentWorksheet(string worksheetName)
        {
            if (_internalDataSet == null)
                throw new ApplicationException("Internal DataSet is null");

            if (!_internalDataSet.Tables.Contains(worksheetName))
                throw new ApplicationException(string.Format(
                    "Worksheet with name {0} does not exist in this Spreasheet", worksheetName));

            _currentActiveWorksheet = _internalDataSet.Tables.IndexOf(worksheetName);
            return this;
        }

        /// <summary>
        /// Sets the current worksheet.
        /// </summary>
        /// <param name="worksheetIndex">Index of the worksheet.</param>
        /// <returns></returns>
        public IExcelReaderProvider SetCurrentWorksheet(int worksheetIndex)
        {
            if (_internalDataSet == null)
                throw new ApplicationException("Internal DataSet is null");

            if (_internalDataSet.Tables.Count <= worksheetIndex)
                throw new ApplicationException(
                    string.Format("Worksheet with index {0} does not exist in this Spreasheet", worksheetIndex));

            _currentActiveWorksheet = worksheetIndex;
            return this;
        }

        /// <summary>
        /// Withes the first name of the row as column.
        /// </summary>
        /// <param name="isFirstRowAsColumnNames">if set to <c>true</c> [is first row as column names].</param>
        /// <returns></returns>
        public IExcelReaderProvider WithFirstRowAsColumnName(bool isFirstRowAsColumnNames)
        {
            _isFirstRowAsColumnNames = isFirstRowAsColumnNames;
            return this;
        }

        public IExcelReaderProvider LoadFromOpenXMLFile(FileStream stream)
        {
            _reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            if (!_reader.IsValid)
            {
                throw new ApplicationException("OpenXML Format not recognized");
            }
            LoadInternalDataSet();
            return this;
        }

        /// <summary>
        /// Toes the data set.
        /// </summary>
        /// <returns></returns>
        public DataSet ToDataSet()
        {
            return _internalDataSet;
        }
    }
}