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

        public object GetValueFromCell(int column, int row)
        {
            throw new NotImplementedException();
        }

        public object GetValueFromCell(string column)
        {
            throw new NotImplementedException();
        }

        public DataTable GetValueFromNamedCell(int columnStart, int rowStart, int colunmnEnd, int rowEnd)
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

        public IExcelReaderProvider LoadFromBinaryFile(FileStream stream)
        {
            _reader = ExcelReaderFactory.CreateBinaryReader(stream);
            if (!_reader.IsValid)
            {
                throw new ApplicationException("BIFF Format not recognized");
            }

            return !_reader.IsValid ? null : this;
        }

        public IExcelReaderProvider LoadFromOpenXMLFile(FileStream stream)
        {
            _reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            if (!_reader.IsValid)
            {
                throw new ApplicationException("OpenXML Format not recognized");
            }

            return !_reader.IsValid ? null : this;
        }

        public DataSet ToDataSet()
        {
            throw new NotImplementedException();
        }

        //public IQueryable GetWorksheetContent(int id)
        //{
        //    var res = _reader.AsDataSet();
        //    return res.Tables[0].AsEnumerable().AsQueryable();
        //}

        //public object GetValue(int row, int column)
        //{
        //    _ds = _reader.AsDataSet();
        //    return _ds.Tables[0].Rows[row][column];
        //}
    }
}