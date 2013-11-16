using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using LinqToExcel;

namespace ProgM.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            var xl = new LinqToExcel.ExcelQueryFactory("C:\\Users\\Danang\\Desktop\\Book1.xlsx");
            var ws = xl.GetWorksheetNames();
            var cl = xl.GetColumnNames(ws.FirstOrDefault());
            var rs = (from c in xl.WorksheetNoHeader(ws.FirstOrDefault())
                                   select c).ToList();

        }
    }
}
