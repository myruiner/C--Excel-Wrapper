﻿// -----------------------------------------------------------------------
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
        object GetValueFromCell(int column, int row);

        object GetValueFromCell(string cellDescription);

        object GetValueFromNamedCell(string namedCell);

        DataTable GetValueFromRange(int columnStart, int rowStart, int colunmnEnd, int rowEnd);

        DataTable GetValueFromRange(string text);

        DataTable GetValueFromNamedRange(string namedRange);

        /// <summary>
        /// Return the named specified Worksheet as a DataTable
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <returns></returns>
        DataTable GetWorksheetContent(string worksheetName);

        /// <summary>
        /// Return the index specified Worksheet as a DataTable
        /// </summary>
        /// <param name="worksheetIndex">Index of the worksheet.</param>
        /// <returns></returns>
        DataTable GetWorksheetContent(int worksheetIndex);

        IExcelReaderProvider LoadFromBinaryFile(FileStream stream);

        IExcelReaderProvider LoadFromOpenXMLFile(FileStream stream);

        DataSet ToDataSet();

        IExcelReaderProvider WithFirstRowAsColumnName(bool isFirstRowAsColumnName);

        IExcelReaderProvider SetCurrentWorksheet(string name);

        IExcelReaderProvider SetCurrentWorksheet(int index);
    }
}