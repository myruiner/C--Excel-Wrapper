// -----------------------------------------------------------------------
// <copyright file="ExcelDataReaderRepository.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System.IO;
using Excel;

namespace Model
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class ExcelDataReaderRepository : IExcelReaderRepository
    {
        private IExcelDataReader _reader;

        public void Dispose()
        {
            if (_reader != null)
            {
                _reader.Dispose();
            }
        }

        public object LoadWorkbook(FileStream filestream)
        {
            return LoadBinaryWorkbook(filestream);
        }

        public object LoadBinaryWorkbook(FileStream filestream)
        {
            _reader = ExcelReaderFactory.CreateBinaryReader(filestream);
            return !_reader.IsValid ? null : _reader;
        }
    }
}