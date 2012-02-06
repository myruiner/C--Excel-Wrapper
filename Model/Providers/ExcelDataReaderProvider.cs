// -----------------------------------------------------------------------
// <copyright file="ExcelDataReaderRepository.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Data;
using System.IO;
using Excel;

namespace Model.Providers
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class ExcelDataReaderProvider : IExcelReaderProvider
    {
        private DataSet _internalDataSet;
        private bool _isFirstRowAsColumnNames;
        private IExcelDataReader _reader;
        private int _currentActiveWorksheet;

        /// <summary>
        /// Gets the name of the value by.
        /// </summary>
        /// <param name="cellName">Name of the cell.</param>
        /// <returns></returns>
        public object GetValueByName(string cellName)
        {
            object res = null;
            while (_reader.Read())
            {
                int index = _reader.GetOrdinal(cellName);
            }
            return res;
        }

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

        public object GetValueFromCell(int column, int row)
        {
            throw new NotImplementedException();
        }

        public object GetValueFromCell(string column)
        {
            throw new ApplicationException("This provider does not support accessing a cell with the corresponding Text description");
        }

        public DataTable GetValueFromRange(int columnStart, int rowStart, int colunmnEnd, int rowEnd)
        {
            throw new NotImplementedException();
        }

        public object GetValueFromNamedCell(string columnStart)
        {
            throw new ApplicationException("This provider does not support named cell");
        }

        public DataTable GetValueFromNamedRange(string name)
        {
            throw new ApplicationException("This provider does not support named range");
        }

        public DataTable GetValueFromRange(string text)
        {
            throw new ApplicationException("This provider does not support accessing a range by the corresponding Text description");
        }

        /// <summary>
        /// Return the named specified Worksheet as a DataTable
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <returns></returns>
        public DataTable GetWorksheetContent(string worksheetName)
        {
            if (_internalDataSet == null)
                throw new ApplicationException("Internal DataSet is null");
            if (!_internalDataSet.Tables.Contains(worksheetName))
                throw new ApplicationException(string.Format("Worksheet with name {0} does not exist in this Spreasheet", worksheetName));

            return _internalDataSet.Tables[worksheetName];
        }

        /// <summary>
        /// Return the index specified Worksheet as a DataTable
        /// </summary>
        /// <param name="worksheetIndex">Index of the worksheet.</param>
        /// <returns></returns>
        public DataTable GetWorksheetContent(int worksheetIndex)
        {
            if (_internalDataSet == null)
                throw new ApplicationException("Internal DataSet is null");
            if (_internalDataSet.Tables.Count <= worksheetIndex)
                throw new ApplicationException(string.Format("Worksheet with index {0} does not exist in this Spreasheet", worksheetIndex));

            return _internalDataSet.Tables[worksheetIndex];
        }

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

        public IExcelReaderProvider SetCurrentWorksheet(string name)
        {
            return this;
        }

        public IExcelReaderProvider SetCurrentWorksheet(int index)
        {
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