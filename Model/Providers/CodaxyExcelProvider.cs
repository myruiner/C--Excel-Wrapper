// -----------------------------------------------------------------------
// <copyright file="CodaxyExcelDataReaderRepository.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Data;
using System.IO;
using Codaxy.Xlio;

namespace Model.Providers
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class CodaxyExcelProvider : IExcelReaderProvider
    {
        private SheetCollection _sheetCollection;

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
        }

        public DataSet GetDataSet()
        {
            throw new NotImplementedException();
        }

        public object GetValueFromCell(int column, int row)
        {
            throw new NotImplementedException();
        }

        public object GetValueFromCell(string column)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromRange(int columnStart, int rowStart, int colunmnEnd, int rowEnd)
        {
            throw new NotImplementedException();
        }

        public object GetValueFromNamedCell(string columnStart)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromNamedRange(string name)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromRange(string text)
        {
            throw new NotImplementedException();
        }

        public DataTable GetWorksheetContent(string worksheetName)
        {
            throw new NotImplementedException();
        }

        public DataTable GetWorksheetContent(int worksheetIndex)
        {
            throw new NotImplementedException();
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
                var wb = Workbook.ReadStream(stream);
                _sheetCollection = wb.Sheets;
            }
            catch
            {
                throw new ApplicationException("OpenXML Format not recognized");
            }

            return this;
        }

        public DataSet ToDataSet(bool isFirstRowAsColumnNames)
        {
            throw new NotImplementedException();
        }

        public object GetValue(int row, int column)
        {
            throw new NotImplementedException();
        }
    }
}