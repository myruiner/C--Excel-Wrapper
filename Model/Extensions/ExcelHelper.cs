// -----------------------------------------------------------------------
// <copyright file="ExcelHelper.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System;
using System.IO;
using Model.Frame;

namespace Model.Extensions
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public static class ExcelHelper
    {
        /// <summary>
        /// Copies the stream to temp.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="extension">The extension.</param>
        /// <returns></returns>
        public static TemporaryFileInfo CopyStreamToTempFileInfo(this Stream source, string extension = null)
        {
            var path = Path.GetTempPath();
            var fileName = Guid.NewGuid().ToString();
            var targetPath = Path.Combine(path, fileName);

            if (!string.IsNullOrEmpty(extension))
                targetPath = targetPath + "." + extension;

            using (var target = File.Open(targetPath, FileMode.OpenOrCreate, FileAccess.Write))
            {
                var buffer = new byte[0x10000];
                try
                {
                    int bytes;
                    while ((bytes = source.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        target.Write(buffer, 0, bytes);
                    }
                }
                finally
                {
                    target.Flush();
                    target.Close();
                }
            }
            return new TemporaryFileInfo(targetPath);
        }
    }
}