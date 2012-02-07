// -----------------------------------------------------------------------
// <copyright file="LinqToExcelProvider.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System.Data;
using System.IO;
using System.Linq;
using LinqToExcel;
using Model.Extensions;

namespace Model.Providers
{
    using System;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class LinqToExcelProvider : IExcelReaderProvider
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

        public DataTable GetWorksheetContent()
        {
            throw new NotImplementedException();
        }

        public IExcelReaderProvider LoadFromBinaryFile(FileStream stream)
        {
            var randomOutput = Path.GetRandomFileName();
            var randomFullPath = Path.Combine(Path.GetTempPath(), randomOutput);

            var tempFile = stream.CopyStreamToTemp();

            var excelReader = new ExcelQueryFactory(tempFile.FullName);

            var indianaCompanies = from c in excelReader.WorksheetRangeNoHeader("A1", "G10") //Selects data within the B3 to G10 cell range
                                   select c;
            tempFile.Delete();
            return this;
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