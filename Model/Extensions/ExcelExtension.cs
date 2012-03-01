// -----------------------------------------------------------------------
// <copyright file="ExcelExtension.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.Text.RegularExpressions;

namespace ExcelReader.Extensions
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public static class ExcelExtension
    {
        public static ExcelCell ToExcelCell(this string columnName)
        {
            var matching = Regex.Match(columnName, "(^[a-zA-Z]+)([1-9]+)$", RegexOptions.IgnoreCase);
            if (!matching.Success || matching.Groups.Count != 3)
                throw new ArgumentException(String.Format("Excel column Name :{0} is invalid :", columnName));

            var col = 0;
            var colName = matching.Groups[1].Value;

            // Process each letter.
            foreach (var t in colName.ToUpperInvariant())
            {
                col *= 26;
                char letter = t;

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                col += letter - 'A' + 1;
            }

            var cell = new ExcelCell { Row = int.Parse(matching.Groups[2].Value) - 1, Column = col - 1 };
            return cell;
        }
    }
}