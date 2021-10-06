using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace EEEEEEEEE
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");


            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"https://portal.petrocollege.ru/Lists/2014/Attachments/10/%D0%B3%D1%80%D1%83%D0%BF%D0%BF_27.09.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
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
                    
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой
        //https://portal.petrocollege.ru/Lists/2014/Attachments/10/%D0%B3%D1%80%D1%83%D0%BF%D0%BF_27.09.xlsx
        }
    }
}
