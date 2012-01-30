// -----------------------------------------------------------------------
// <copyright file="CodaxyExcelDataReaderRepository.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Data;
using System.IO;
using System.Linq;
using Codaxy.Xlio;

namespace Model.CodaxyRepository
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class CodaxyExcelDataReaderRepository : IExcelReaderRepository
    {
        private SheetCollection _sheetCollection;

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
        }

        /// <summary>
        /// Loads the binary workbook.
        /// </summary>
        /// <param name="filestream">The filestream.</param>
        /// <returns></returns>
        public IExcelReaderRepository LoadBinaryWorkbook(FileStream filestream)
        {
            throw new ApplicationException("Binary Format not supported by CodaxyExcelDataReader");
        }

        /// <summary>
        /// Loads the open XML workbook.
        /// </summary>
        /// <param name="filestream">The filestream.</param>
        /// <returns></returns>
        public IExcelReaderRepository LoadOpenXmlWorkbook(FileStream filestream)
        {
            try
            {
                var wb = Workbook.ReadStream(filestream);
                _sheetCollection = wb.Sheets;
            }
            catch
            {
                throw new ApplicationException("OpenXML Format not recognized");
            }

            return this;
        }

        public DataSet GetDataSet()
        {
            throw new NotImplementedException();
        }

        public IQueryable GetWorksheetContent(int id)
        {
            throw new NotImplementedException();
        }

        public object GetValue(int row, int column)
        {
            throw new NotImplementedException();
        }
    }
}