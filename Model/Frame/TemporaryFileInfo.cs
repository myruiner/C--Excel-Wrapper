// -----------------------------------------------------------------------
// <copyright file="TemporaryFileInfo.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace Model.Frame
{
    using System;
    using System.IO;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class TemporaryFileInfo : IDisposable
    {
        private readonly FileInfo _internalFileInfo;

        /// <summary>
        /// Initializes a new instance of the <see cref="TemporaryFileInfo"/> class.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        public TemporaryFileInfo(string fileName)
        {
            _internalFileInfo = new FileInfo(fileName);
        }

        /// <summary>
        /// Return current FileInfo Object
        /// </summary>
        /// <returns></returns>
        public FileInfo GetFileFinfo()
        {
            return _internalFileInfo;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_internalFileInfo != null)
            {
                _internalFileInfo.Delete();
            }
        }
    }
}