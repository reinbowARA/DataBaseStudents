using Aspose.Cells;
using static System.Console;

Workbook wb = new Workbook("Students.xls"); // открыть существующий xls файл
int n = Convert.ToInt32(Console.ReadLine());
//Add new worksheet in existing workbook
Worksheet sheet = wb.Worksheets.Add($"Groop {n}");

//Access the "A1" cell in the sheet.
Cell cell = sheet.Cells["A1"];

//Input the "Hello World!" text into the "A1" cell
cell.PutValue("Hello World!");
for (int i = 0; i <= n; i++)
{
    if (i == 0)
    {
        wb.Worksheets.RemoveAt("Evaluation Warning");
    }
    wb.Worksheets.RemoveAt($"Evaluation Warning ({i})");
}
Aspose.Cells.License license = new Aspose.Cells.License();
license.SetLicense("Aspose.Cells.lic");
//Save the Excel file.
wb.Save("Students.xls");
/*
сделать проверку наличия ексель файла, если нету создать
сделать меню с переходами на создать группу добавить студента и их удаления
вывод в консоль содержимое группы
*/