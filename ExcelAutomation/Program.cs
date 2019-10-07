using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAutomation
{
    class Program
    {
        private static int _done = 0;
        private const uint FirstRow = 5;
        private const uint LastRow = 915;
        static void Main(string[] args)
        {
            var oExcelApp = new Excel.Application {Visible = true};
            //Get reference to Excel.Application from the ROT.
            //oExcelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            oExcelApp.Workbooks.Open(@"C:\Dat\control group 911.xlsx");
            //Display the name of the object.zx
            Console.WriteLine("Active Workbook:");
            Console.WriteLine(oExcelApp.ActiveWorkbook.Name);
            var activeSheet = (Excel._Worksheet)oExcelApp.ActiveWorkbook.ActiveSheet;
            Console.WriteLine("Transforming");
            Parallel.For(FirstRow, LastRow + 1, i =>
              {
                  int lesionColumnIndex = 'J' - 'A' + 1;
                  string lesions = ((Excel.Range)activeSheet.Cells[i, lesionColumnIndex]).Value2?.ToString();
                  //for (uint q = 1; q <= 9; q++)
                  //{
                  Parallel.For(1, 10, q =>
                  {
                      if (lesions == null) activeSheet.Cells[i, lesionColumnIndex + q] = 0;
                      else activeSheet.Cells[i, lesionColumnIndex + q] = (lesions.Contains(q.ToString())) ? 1 : 0;
                  });
                  //}

                  int symptomColumnIndex = 'W' - 'A' + 1;
                  string symptomps = ((Excel.Range)activeSheet.Cells[i, symptomColumnIndex]).Value2?.ToString();
                  //for (uint q = 1; q <= 11; q++)
                  //{
                  Parallel.For(1, 12, q =>
                  {
                      if (symptomps == null) activeSheet.Cells[i, symptomColumnIndex + q] = 0;
                      else activeSheet.Cells[i, symptomColumnIndex + q] = (symptomps.Contains(q.ToString())) ? 1 : 0;
                  });
                  //}

                  Interlocked.Increment(ref _done);
                  DisplayProgress();
              });
            Console.WriteLine("");
            Console.WriteLine("Complete! Press Any Key To Exit");
            Console.ReadLine();
        }

        static void DisplayProgress()
        {
            var currentCursorPosition = Console.CursorTop;
            Console.SetCursorPosition(0, currentCursorPosition);
            // Clear Line
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0,currentCursorPosition);
            Console.Write($"{Math.Round(_done * 100.0 / (LastRow - FirstRow), 2)}% : {_done} / {(LastRow - FirstRow)}");
        }
    }
}
