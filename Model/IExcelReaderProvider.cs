// -----------------------------------------------------------------------
// <copyright file="IExcelReaderProvider.cs" company="Trivadis AG">
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
    public interface IExcelReaderProvider : IDisposable
    {
        IExcelReaderProvider ConfigureProvider();

        object GetValueFromCell(int column, int row);

        object GetValueFromCell(string column);

        DataTable GetValueFromNamedCell(int columnStart, int rowStart, int colunmnEnd, int rowEnd);

        object GetValueFromNamedCell(string columnStart);

        DataTable GetValueFromNamedRange(string name);

        DataTable GetValueFromRange(string text);

        DataTable GetWorksheetContent(string worksheetName);

        DataTable GetWorksheetContent(int worksheetIndex);

        IExcelReaderProvider LoadFromBinaryFile(FileStream stream);

        void LoadFromFile(FileStream stream);

        IExcelReaderProvider LoadFromOpenXMLFile(FileStream stream);

        DataSet ToDataSet();
    }
}