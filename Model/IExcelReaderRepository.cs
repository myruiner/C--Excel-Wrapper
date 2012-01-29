using System;
using System.IO;

namespace Model
{
    public interface IExcelReaderRepository : IDisposable
    {
        Object LoadWorkbook(FileStream filestream);
    }
}