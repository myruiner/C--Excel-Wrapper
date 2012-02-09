// -----------------------------------------------------------------------
// <copyright file="IExcelReaderProvider.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Collections;
using System.Data;
using System.IO;

namespace ExcelReader
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public interface IExcelReaderProvider : IDisposable
    {
        /// <summary>
        /// Gets the content of cell identified by his ID
        /// </summary>
        /// <param name="address"> </param>
        /// <returns></returns>
        object GetValueFromCellByAddress(string address);

        /// <summary>
        /// Gets the content of cell identified by ID
        /// </summary>
        /// <param name="column">The column.</param>
        /// <param name="row">The row.</param>
        /// <returns></returns>
        object GetValueFromCellByID(int column, int row);

        /// <summary>
        /// Gets the content of cell identified by his name
        /// </summary>
        /// <param name="namedCell">The named cell.</param>
        /// <returns></returns>
        object GetValueFromCellByName(string namedCell);

        /// <summary>
        /// Gets the content of a range identified by his ID
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="address"> </param>
        /// <returns></returns>
        DataTable GetValueFromRangeByAddress(string address);

        /// <summary>
        /// Gets the content of a range identified by his ID
        /// </summary>
        /// <param name="columnStart">The column start.</param>
        /// <param name="rowStart">The row start.</param>
        /// <param name="colunmnEnd">The colunmn end.</param>
        /// <param name="rowEnd">The row end.</param>
        /// <returns></returns>
        DataTable GetValueFromRangeByID(int columnStart, int rowStart, int colunmnEnd, int rowEnd);

        /// <summary>
        /// Gets the content of a range identified by his name
        /// </summary>
        /// <param name="namedRange">The named range.</param>
        /// <returns></returns>
        DataTable GetValueFromRangeByName(string namedRange);

        /// <summary>
        /// Return the current speicified Worksheet as a DataTable
        /// </summary>
        /// <returns></returns>
        IEnumerable GetWorksheetContent();

        /// <summary>
        /// Loads Content from a Binary file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        IExcelReaderProvider LoadFromBinaryFile(FileStream stream);

        /// <summary>
        /// Loads content from an OpenXML file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        IExcelReaderProvider LoadFromOpenXMLFile(FileStream stream);

        /// <summary>
        /// Sets the current worksheet by Name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        IExcelReaderProvider SetCurrentWorksheet(string name);

        /// <summary>
        /// Sets the current worksheet by his index
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        IExcelReaderProvider SetCurrentWorksheet(int index);

        /// <summary>
        /// Toes the data set.
        /// </summary>
        /// <returns></returns>
        DataSet ToDataSet();

        /// <summary>
        /// Withes the first Rows containing the Column names.
        /// </summary>
        /// <param name="isFirstRowAsColumnName">if set to <c>true</c> [is first row as column name].</param>
        /// <returns></returns>
        IExcelReaderProvider WithFirstRowAsColumnName(bool isFirstRowAsColumnName);
    }
}