using System;
using System.IO;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;

namespace EEEEEEEEE
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            string url = @"https://portal.petrocollege.ru/Lists/2014/Attachments/10/%D0%B3%D1%80%D1%83%D0%BF%D0%BF_27.09.xlsx";
            
            using WebClient client = new()
            {
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential("10190997", "2KU3bmZWyC"),
            };
            client.Headers.Add(HttpRequestHeader.Cookie, null);
            client.DownloadFile(url, "../../../resources/timetable.xlsx");
            client.Dispose(); // ВАЖНО!! закрывает процесс экселя. иначе будет конфликт процессов

            string filePath = Path.GetFullPath("../../../resources/timetable.xlsx"); // получает путь к файлу

            Excel.Application ObjWorkExcel = new(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filePath); //открыть файл 
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
            for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
            {
                Console.Write("\n");
                for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                {
                    list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку
                    Console.Write(list[i, j]);
                }
            }

            ObjWorkBook.Close(false); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой

            // уничтожает к чертям абсолютно все процессы экселя. very opasno
            System.Diagnostics.Process[] localByName = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process pr in localByName)
                if (pr.ProcessName == "EXCEL") pr.Kill();


            // https://portal.petrocollege.ru/Lists/2014/Attachments/10/%D0%B3%D1%80%D1%83%D0%BF%D0%BF_27.09.xlsx
        }
    }
}
