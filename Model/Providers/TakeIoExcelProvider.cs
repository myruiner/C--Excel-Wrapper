// -----------------------------------------------------------------------
// <copyright file="TakeIoExcelProvider.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Collections;
using System.Data;
using System.IO;

namespace ExcelReader.Providers
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class TakeIoExcelProvider : IExcelReaderProvider
    {
        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public object GetValueFromCellByID(int column, int row)
        {
            throw new NotImplementedException();
        }

        public object GetValueFromCellByAddress(string address)
        {
            throw new NotImplementedException();
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
            throw new NotImplementedException();
        }

        public IExcelReaderProvider LoadFromBinaryFile(FileStream stream)
        {
            throw new NotImplementedException();
        }

        public IExcelReaderProvider LoadFromOpenXMLFile(FileStream stream)
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