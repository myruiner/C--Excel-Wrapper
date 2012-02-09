// -----------------------------------------------------------------------
// <copyright file="ExcelReaderSetup.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.IO;

namespace ExcelReader
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class ExcelReader : IDisposable
    {
        private readonly IExcelReaderProvider _provider;

        /// <summary>
        /// Prevents a default instance of the <see cref="ExcelReader"/> class from being created.
        /// </summary>
        /// <param name="provider">The provider.</param>
        private ExcelReader(IExcelReaderProvider provider)
        {
            _provider = provider;
        }

        /// <summary>
        /// Loads the file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="isOpenXML"> </param>
        /// <returns></returns>
        public IExcelReaderProvider LoadFile(FileStream stream, bool isOpenXML)
        {
            if (isOpenXML)
                _provider.LoadFromOpenXMLFile(stream);
            else
                _provider.LoadFromBinaryFile(stream);
            return _provider;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_provider != null)
                _provider.Dispose();
        }

        /// <summary>
        /// Configures the provider.
        /// </summary>
        /// <param name="provider">The provider.</param>
        /// <returns></returns>
        public static ExcelReader ConfigureProvider(IExcelReaderProvider provider)
        {
            return new ExcelReader(provider);
        }
    }
}