// -----------------------------------------------------------------------
// <copyright file="CodaxyExcelDataReaderRepository.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Collections;
using System.Data;
using System.IO;
using Codaxy.Xlio;
using ExcelReader.Extensions;

namespace ExcelReader.Providers
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class CodaxyExcelProvider : IExcelReaderProvider
    {
        private string _currentWorksheetName;
        private SheetCollection _sheetCollection;

        /// <summary>
        /// Gets or sets the name of the current worksheet.
        /// </summary>
        /// <value>
        /// The name of the current worksheet.
        /// </value>
        public string CurrentWorksheetName
        {
            get
            {
                if (string.IsNullOrEmpty(_currentWorksheetName))
                    _currentWorksheetName = _sheetCollection[0].SheetName;
                return _currentWorksheetName;
            }
            set { _currentWorksheetName = value; }
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
        }

        /// <summary>
        /// Gets the content of cell identified by ID
        /// </summary>
        /// <param name="column">The column.</param>
        /// <param name="row">The row.</param>
        /// <returns></returns>
        public object GetValueFromCellByID(int column, int row)
        {
            return _sheetCollection[CurrentWorksheetName].Cells[row, column].Value ?? DBNull.Value;
        }

        /// <summary>
        /// Gets the value from cell by address.
        /// </summary>
        /// <param name="cellAddress">The cell address.</param>
        /// <returns></returns>
        public object GetValueFromCellByAddress(string cellAddress)
        {
            ExcelCell cell = cellAddress.ToExcelCell();
            return GetValueFromCellByID(cell.Column, cell.Row);
        }

        public object GetValueFromCellByName(string columnStart)
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
            var dataTable = new DataTable();
            return dataTable.AsEnumerable();
        }

        /// <summary>
        /// Loads from binary file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public IExcelReaderProvider LoadFromBinaryFile(FileStream stream)
        {
            throw new ApplicationException("Binary Format not supported by CodaxyExcelDataReader");
        }

        /// <summary>
        /// Loads from open XML file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public IExcelReaderProvider LoadFromOpenXMLFile(FileStream stream)
        {
            try
            {
                Workbook wb = Workbook.ReadStream(stream);
                _sheetCollection = wb.Sheets;
            }
            catch
            {
                throw new ApplicationException("OpenXML Format not recognized");
            }

            return this;
        }

        public DataSet ToDataSet()
        {
            throw new ApplicationException("Methode is not supported by this provider");
        }

        public IExcelReaderProvider WithFirstRowAsColumnName(bool isFirstRowAsColumnName)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Sets the current worksheet by Name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public IExcelReaderProvider SetCurrentWorksheet(string name)
        {
            if (_sheetCollection[name] != null)
                CurrentWorksheetName = name;
            return this;
        }

        /// <summary>
        /// Sets the current worksheet by his index
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public IExcelReaderProvider SetCurrentWorksheet(int index)
        {
            if (_sheetCollection[index] != null)
                CurrentWorksheetName = _sheetCollection[index].SheetName;
            return this;
        }
    }
}