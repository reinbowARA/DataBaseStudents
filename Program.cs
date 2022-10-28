using Aspose.Cells;
using static System.Console;

// проверить или создать xls файл
Workbook wb = new Workbook("Students.xls");
int countSheet = wb.Worksheets.Count;
int cheek = 0;
int n;
do
{
    WriteLine("------MENU------");
    WriteLine("1 - Создать группу");
    WriteLine("2 - Удалить группу");
    WriteLine("3 - Список групп");
    WriteLine("4 - Добавить студента/студентов в группу");
    WriteLine("5 - Состав группы");
    WriteLine("6 - Удаление студента из группы");
    WriteLine("0 - exit");
    
    n = Convert.ToInt32(ReadLine());
    switch (n)
    {
        case 1:
            String? GroopName = ReadLine();
            for (int i = 0; i < countSheet; i++)
            {
                if(GroopName == wb.Worksheets[i].Name)
                {
                    WriteLine("Такая группа уже существует");
                    cheek++;
                    break;
                }
            }
            if (cheek == 0)
            {
                Worksheet sheet = wb.Worksheets.Add(GroopName);
                WriteLine("Группа создана");
            }
            wb.Save("Students.xls");
            break;
        case 2:
            GroopName = ReadLine();
            for (int i = 0; i < countSheet; i++)
            {
                if (GroopName == wb.Worksheets[i].Name)
                {
                    wb.Worksheets.RemoveAt(GroopName);
                    WriteLine("Группа удаленна");
                    cheek++;
                    break;
                }
            }               
            if (cheek == 0)
            {
                WriteLine("нет такой группы");
            }  
            wb.Save("Students.xls");         
            break;
        case 3:
            for (int i = 0; i < countSheet; i++)
            {
                WriteLine(wb.Worksheets[i].Name);
            }
            break;
        case 4:
            GroopName = ReadLine();
            Cell cell;
            String? NameStudent;
            for (int j = 0; j < 30; j++)
            {
                Write($"{j} - Введите имя студента: ");
                NameStudent = ReadLine();
                cell = wb.Worksheets[GroopName].Cells[j,0];
                cell.PutValue(NameStudent);
                if (NameStudent?.Length <= 0)
                {
                    WriteLine("Запись закончена");
                    break;
                }
            }
            wb.Save("Students.xls");
            break;
        case 5:
            GroopName = ReadLine();
            AutoSort(GroopName);
            for (int j = 0; j < 30; j++)
            {
                cell = wb.Worksheets[GroopName].Cells[j,0];
                WriteLine((j+1)+" - "+cell.StringValue);
                if(cell.StringValue.Length == 0){
                    break;
                }
            }
            wb.Save("Students.xls");
            break;
        case 6:
            GroopName = ReadLine();
            NameStudent = ReadLine();
            for (int j = 0; j < 30; j++)
            {
                cell = wb.Worksheets[GroopName].Cells[j,0];
                if (NameStudent == cell.StringValue)
                {
                    cell.PutValue("");
                }
                if(cell.StringValue.Length == 0){
                    break;
                }
            }
            AutoSort(GroopName);
            wb.Worksheets[GroopName].Cells.DeleteBlankRows();
            wb.Save("Students.xls");
            break;
        default:
            break;
    }
} while (n != 0);

for (int i = 0; i <= wb.Worksheets.Count; i++)
{
    if (i == 0)
    {
        wb.Worksheets.RemoveAt("Evaluation Warning");
    }
    wb.Worksheets.RemoveAt($"Evaluation Warning ({i})");
}
wb.Save("Students.xls");

static void AutoSort(string? GroopName){
    var wb = new Workbook("Students.xls");
    DataSorter sorter = wb.DataSorter;
    sorter.Order1 = SortOrder.Ascending;
    sorter.Key1 = 0;
    sorter.Order2 =SortOrder.Descending;
    sorter.Key2 = 1;
    CellArea ca = new CellArea();
    ca.StartRow = 0;
    ca.StartColumn = 0;
    ca.EndRow = 29;
    ca.EndColumn = 0;
    sorter.Sort(wb.Worksheets[GroopName].Cells, ca);
    wb.Worksheets[GroopName].Cells.DeleteBlankRows();
}