using System;
using System.Data;
using System.IO;
using System.Linq;

namespace Model
{
    public interface IExcelReaderRepository : IDisposable
    {
        IExcelReaderRepository LoadBinaryWorkbook(FileStream filestream);

        IExcelReaderRepository LoadOpenXmlWorkbook(FileStream filestream);

        DataSet GetDataSet();

        IQueryable GetWorksheetContent(int id);

        object GetValue(int row, int column);
    }
}