using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Service
{
    using System;
    using Model;

    public sealed class MyConsts
    {
        private MyConsts() { }

        public static readonly string Val1 = "MyVal1";
        public static readonly string Val2 = "MyVal2";
        public static readonly string Val3 = "MyVal3";
    }

    public class ExcelService
    {
        private IExcelReaderRepository _repository;

        public ExcelService(IExcelReaderRepository repository)
        {
            {
                _repository = repository;
                if (repository == null)
                {
                    throw new InvalidOperationException("Repository can not be null");
                }
            }
        }

        public ExcelService LoadFile(FileStream fileStream)
        {
           _repository.
            return this;
        }

        public IList<Worksheet> GetWorksheets()
        {
            throw new Exception("not implemented");
        }

        public DataSet GetDataSet()
        {
            throw new Exception("not implemented");
        }
    }
}