// -----------------------------------------------------------------------
// <copyright file="ExcelDataReaderRepository.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Data;
using System.IO;
using System.Linq;
using Excel;

namespace Model.ExcelDataRepository
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class ExcelDataReaderRepository : IExcelReaderRepository
    {
        private IExcelDataReader _reader;

        private DataSet _ds;

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
        /// Loads the open XML workbook.
        /// </summary>
        /// <param name="filestream">The filestream.</param>
        /// <returns></returns>
        public IExcelReaderRepository LoadOpenXmlWorkbook(FileStream filestream)
        {
            _reader = ExcelReaderFactory.CreateOpenXmlReader(filestream);
            if (!_reader.IsValid)
            {
                throw new ApplicationException("OpenXML Format not recognized");
            }

            return !_reader.IsValid ? null : this;
        }

        /// <summary>
        /// Loads the binary workbook.
        /// </summary>
        /// <param name="filestream">The filestream.</param>
        /// <returns></returns>
        public IExcelReaderRepository LoadBinaryWorkbook(FileStream filestream)
        {
            _reader = ExcelReaderFactory.CreateBinaryReader(filestream);
            if (!_reader.IsValid)
            {
                throw new ApplicationException("BIFF Format not recognized");
            }

            return !_reader.IsValid ? null : this;
        }

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
                var index = _reader.GetOrdinal(cellName);
            }
            return res;
        }

        public DataSet GetDataSet()
        {
            var res = _reader.AsDataSet();
            return res;
        }

        public IQueryable GetWorksheetContent(int id)
        {
            var res = _reader.AsDataSet();
            return res.Tables[0].AsEnumerable().AsQueryable();
        }

        public object GetValue(int row, int column)
        {
            _ds = _reader.AsDataSet();
            return _ds.Tables[0].Rows[row][column];
        }
    }
}