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

        public object GetValueFromCell(int column, int row)
        {
            throw new NotImplementedException();
        }

        public object GetValueFromCell(string cellDescription)
        {
            throw new NotImplementedException();
        }

        public object GetValueFromNamedCell(string namedCell)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromRange(int columnStart, int rowStart, int colunmnEnd, int rowEnd)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromRange(string text)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromNamedRange(string namedRange)
        {
            throw new NotImplementedException();
        }

        public DataTable GetWorksheetContent()
        {
            throw new NotImplementedException();
        }

        public IExcelReaderProvider LoadFromBinaryFile(FileStream stream)
        {
            throw new NotImplementedException();
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

        public DataSet ToDataSet()
        {
            throw new NotImplementedException();
        }

        public IExcelReaderProvider WithFirstRowAsColumnName(bool isFirstRowAsColumnName)
        {
            throw new NotImplementedException();
        }

        public IExcelReaderProvider SetCurrentWorksheet(string name)
        {
            throw new NotImplementedException();
        }

        public IExcelReaderProvider SetCurrentWorksheet(int index)
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}