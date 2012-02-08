using System.Linq;
using LinqToExcel;
using LinqToExcel.Domain;

namespace ConsoleApplication1
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var excelQueryFactory = new ExcelQueryFactory(@"C:\Users\pth\Documents\Visual Studio 2010\Projects\C--Excel-Wrapper\C--Excel-Wrapper\Tests\ExcelTemplates\Basic.xlsx");
            excelQueryFactory.DatabaseEngine = DatabaseEngine.Ace;

            var query = from c in excelQueryFactory.WorksheetNoHeader("ws2")
                        select c[0];

            var res = query.ToList();
        }
    }
}