// -----------------------------------------------------------------------
// <copyright file="ExcelReaderSetup.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Data;
using System.IO;

namespace Model
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class ExcelReader : IDisposable
    {
        private readonly IExcelReaderProvider _provider;

        /// <summary>
        /// Prevents a default instance of the <see cref="ExcelReader"/> class from being created.
        /// </summary>
        /// <param name="provider">The provider.</param>
        private ExcelReader(IExcelReaderProvider provider)
        {
            _provider = provider;
        }

        /// <summary>
        /// Retreive the content of a given cell
        /// </summary>
        /// <param name="column">The column.</param>
        /// <param name="row">The row.</param>
        /// <returns></returns>
        public object GetValueFromCell(int column, int row)
        {
            return _provider.GetValueFromCell(column, row);
        }

        /// <summary>
        /// Retreive the content of a given cell
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        public object GetValueFromCell(string text)
        {
            return _provider.GetValueFromCell(text);
        }

        /// <summary>
        /// Retreive the content of a given named cell.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public object GetValueFromNamedCell(string name)
        {
            return _provider.GetValueFromNamedCell(name);
        }

        /// <summary>
        /// Retreive the content of a given named range.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public DataTable GetValueFromNamedRange(string name)
        {
            return _provider.GetValueFromNamedRange(name);
        }

        /// <summary>
        /// Gets the value from range.
        /// </summary>
        /// <param name="columnStart">The column start.</param>
        /// <param name="rowStart">The row start.</param>
        /// <param name="colunmnEnd">The colunmn end.</param>
        /// <param name="rowEnd">The row end.</param>
        /// <returns></returns>
        public DataTable GetValueFromRange(int columnStart, int rowStart, int colunmnEnd, int rowEnd)
        {
            return _provider.GetValueFromRange(columnStart, rowStart, colunmnEnd, rowEnd);
        }

        /// <summary>
        /// Gets the value from range.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        public DataTable GetValueFromRange(string text)
        {
            return _provider.GetValueFromRange(text);
        }

        /// <summary>
        /// Gets the content of the worksheet.
        /// </summary>
        /// <param name="worksheetIndex">Index of the worksheet.</param>
        /// <returns></returns>
        public DataTable GetWorksheetContent(int worksheetIndex)
        {
            return _provider.GetWorksheetContent(worksheetIndex);
        }

        /// <summary>
        /// Gets the content of the worksheet.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <returns></returns>
        public DataTable GetWorksheetContent(string worksheetName)
        {
            return _provider.GetWorksheetContent(worksheetName);
        }

        /// <summary>
        /// Loads the file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="isOpenXML"> </param>
        /// <returns></returns>
        public ExcelReader LoadFile(FileStream stream, bool isOpenXML)
        {
            if (isOpenXML)
                _provider.LoadFromOpenXMLFile(stream);
            else
                _provider.LoadFromBinaryFile(stream);
            return this;
        }

        /// <summary>
        /// Toes the data set.
        /// </summary>
        /// <returns></returns>
        public DataSet ToDataSet()
        {
            return _provider.ToDataSet();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_provider != null)
                _provider.Dispose();
        }

        /// <summary>
        /// Configures the provider.
        /// </summary>
        /// <param name="provider">The provider.</param>
        /// <returns></returns>
        public static ExcelReader ConfigureProvider(IExcelReaderProvider provider)
        {
            return new ExcelReader(provider);
        }
    }
}