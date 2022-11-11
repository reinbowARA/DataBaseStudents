using Aspose.Cells;
using static System.Console;

try{FileStream fs = File.Open("Students.xls",FileMode.OpenOrCreate);fs?.Close();}catch(Exception ex){WriteLine(ex.ToString());}
String NameFileXLS = "Students.xls";
Workbook wb = new Workbook(NameFileXLS);
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
    WriteLine("5 - Состав группы / потока");
    WriteLine("6 - Удаление студента из группы");
    WriteLine("7 - Создание потока");
    WriteLine("8 - сортировка по ФИО/Группам");
    WriteLine("9 - Save in exl");
    WriteLine("0 - exit");
    

    n = Convert.ToInt32(ReadLine());
    switch (n)
    {
        case 1:
            Write("Введите название группы вида  ####-##-##: ");
            String GroopName = ReadLine();
            GroopName.ToUpper();
            try{wb.Worksheets.Add($"{GroopName: ####-##-##}");}catch(Aspose.Cells.CellsException){WriteLine("Группа уже существует");break;}
            WriteLine("Группа создана");
            break;
        case 2:
            Write("Введите полное название группы: ");
            GroopName = ReadLine();
            if (GroopName == wb.Worksheets[GroopName].Name)
            {
                wb.Worksheets.RemoveAt(GroopName);
                WriteLine("Группа удаленна");
                cheek++;
                break;
            }else{
                WriteLine("Данной группы нету в списке");
                break;
            }                     
        case 3:
            try{
                WriteLine("Список групп / потоков");
                for (int i = 0; i < countSheet; i++)
                {
                    string frame = string.Concat(Enumerable.Repeat("-",wb.Worksheets[i].Name.Length));
                    WriteLine("/"+ frame +"\\");
                    WriteLine("|"+wb.Worksheets[i].Name+"|");
                    WriteLine("\\"+frame+"/");
                }
            }catch(System.ArgumentOutOfRangeException){WriteLine("Ошибка: выход индекса за предела");}
            catch(System.TypeInitializationException){WriteLine("Непонятная ошибка");}
            break;
        case 4:
            Write("Введите название группы: ");
            GroopName = ReadLine();
            Cell cell;
            String NameStudent;
            for (int j = 0; j < 30; j++)
            {
                Write($"{j+1} - Введите имя студента: ");
                NameStudent = ReadLine();
                cell = wb.Worksheets[GroopName].Cells[j,0];
                cell.PutValue(NameStudent);
                if (NameStudent?.Length <= 0)
                {
                    WriteLine("Запись закончена");
                    break;
                }
            }
            break;
        case 5:
            Write("Введите название группы / потока: ");
            GroopName = ReadLine();
            for (int j = 0; j <= wb.Worksheets[GroopName].Cells.MaxDataRow; j++)
            {   if(wb.Worksheets[GroopName].Cells.MaxDataRow > 30){
                    var cell1 = wb.Worksheets[GroopName].Cells[j,0];
                    var cell2 = wb.Worksheets[GroopName].Cells[j,1];
                    string frame1 = string.Concat(Enumerable.Repeat("-",cell1.StringValue.Length));
                    string frame2 = string.Concat(Enumerable.Repeat("-",cell2.StringValue.Length));
                    string frame3 = string.Concat(Enumerable.Repeat("-",Convert.ToString(j).Length));
                    WriteLine($"{frame3}\\/{frame1}\\/{frame2}\\");
                    WriteLine($"{(j+1)}|{cell1.StringValue}  |{cell2.StringValue,10}|");
                    WriteLine($"{frame3}/\\{frame1}/\\{frame2}/");
                    if(cell1.StringValue.Length == 0){
                        break;
                    }
                }else{
                    
                    cell = wb.Worksheets[GroopName].Cells[j,0];
                    string frame = string.Concat(Enumerable.Repeat("-",cell.StringValue.Length));
                    WriteLine($"--\\/{frame}\\");
                    WriteLine((j+1)+"|"+cell.StringValue);
                    WriteLine($"--/\\{frame}/");
                    if(cell.StringValue.Length == 0){
                        break;
                    }
                }
            }
            break;
        case 6:
            Write("Введите название группы: ");
            GroopName = ReadLine();
            Write("Введите ФИО студента: ");
            NameStudent = ReadLine();
            for (int j = 0; j < 30; j++)
            {
                cell = wb.Worksheets[GroopName].Cells[j,0];
                if (NameStudent == cell.StringValue)
                {
                    cell.PutValue("");
                }
                if(cell.StringValue.Length == 0 || cell.StringValue == null){
                    break;
                }
            }
            AutoSortName(GroopName,30);
            wb.Worksheets[GroopName].Cells.DeleteBlankRows();
           // wb.Save(NameFileXLS);
            break;
        case 7:
            Write("Введите название потока без года: ");
            String ThreadName = ReadLine();
            Write("Введите год потока: ");
            String EarsThread = ReadLine();
            String thread = ThreadName+"-"+EarsThread;
            int count = 0;
            int addcell = 0;
            try{wb.Worksheets.Add(thread);}catch(Aspose.Cells.CellsException){WriteLine("Данный поток уже существует");}
            for (int i = 0; i < countSheet; i++)
            {
                if(wb.Worksheets[i].Name.StartsWith(ThreadName) && wb.Worksheets[i].Name.EndsWith(EarsThread)){count++;}
            }
            WriteLine(count);
            for (int i = 0; i < countSheet; i++)
            {
                if(wb.Worksheets[i].Name.StartsWith(ThreadName) && wb.Worksheets[i].Name.EndsWith(EarsThread) && (wb.Worksheets[i].Name.Length == 10))
                {
                    String name = wb.Worksheets[i].Name;
                    for (int j = 0; j <= wb.Worksheets[i].Cells.MaxDataRow; j++)
                    {
                        var cellname = wb.Worksheets[name].Cells[j,0].StringValue;
                        wb.Worksheets[thread].Cells[j+addcell,0].PutValue(cellname); 
                        wb.Worksheets[thread].Cells[j+addcell,1].PutValue(name);
                    }
                    addcell = addcell + wb.Worksheets[i].Cells.MaxDataRow+1;
                }
            }            
            break;
        case 8:
            GroopName = ReadLine();
            int countrow = wb.Worksheets[GroopName].Cells.MaxDataRow;
            Write("Вы желаете отсортировать поток по Алфваиту ФИО или по группам (1 - по ФИО, 2 - по группам, enter - оставить без изменений): ");
            int s = Int32.Parse(ReadLine());
            if(s == 1){AutoSortName(GroopName,countrow);}else if(s == 2){AutoSortGrooup(GroopName,countrow);}else{WriteLine("ну лан");}
            break;
        case 9:
            try{wb.Save(NameFileXLS);}catch(Exception){WriteLine("Error save");}
            break;
        default:
            break;
    }

} while (n != 0);

//try{DeleteWarning();}catch(Exception ex){WriteLine(ex.ToString() + "Error here 127");};// он чет здесь ругается
//wb.Save(NameFileXLS);

void AutoSortName(string GroopName,int MaxRow){
   // var wb = new Workbook();
    DataSorter sorter = wb.DataSorter;
    sorter.Order1 = SortOrder.Ascending;
    sorter.Key1 = 0;
    sorter.Order2 = SortOrder.Descending;
    sorter.Key2 = 1;
    CellArea ca = new CellArea();
    ca.StartRow = 0;// start in first row
    ca.StartColumn = 0;// start in first column
    ca.EndRow = MaxRow; // end in 30+ Row
    ca.EndColumn = 1;//как были в одной колонке так и остались
    sorter.Sort(wb.Worksheets[GroopName].Cells, ca);
    wb.Worksheets[GroopName].Cells.DeleteBlankRows();
}
void AutoSortGrooup(string GroopName, int MaxRow){
    DataSorter sorter = wb.DataSorter;
    sorter.Order1 = SortOrder.Ascending;
    sorter.Key1 = 1;
    sorter.Order2 = SortOrder.Ascending;
    sorter.Key2 = 0;
    CellArea ca = new CellArea();
    ca.StartRow = 0;
    ca.StartColumn = 0;
    ca.EndRow = MaxRow;
    ca.EndColumn = 1;
    sorter.Sort(wb.Worksheets[GroopName].Cells, ca);
    wb.Worksheets[GroopName].Cells.DeleteBlankRows();
}
//public static String Repeat(String value,int count) => string.Concat(Enumerable.Repeat(value,count));