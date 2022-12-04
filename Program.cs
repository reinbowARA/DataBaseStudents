using Aspose.Cells;
using static System.Console;

try{FileStream fs = File.Open("Students.xls",FileMode.OpenOrCreate);fs?.Close();}catch(Exception ex){WriteLine(ex.ToString());}
String NameFileXLS = "Students.xls";
Workbook wb = new Workbook(NameFileXLS);
int countSheet = wb.Worksheets.Count;
int cheek = 0;
int idgroop = 0; 
int n;

Menu:
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
ret:  
    Write("Введите число: ");
   try
   {
        n = Convert.ToInt32(ReadLine());
        if(n > 9){WriteLine("Ошибка ввода");goto ret;}
   }
   catch (Exception)
   {
        WriteLine("Неверный формат ввода");
        goto ret;
   }
do
{
    switch (n)
    {
        case 1:
            Write("Введите название группы вида  ББББ-ЦЦ-ЦЦ (Б - буква, Ц - цифра): ");
            String GroopName = ReadLine();
            GroopName.ToUpper();
            if(!ChekerGroop(GroopName))
            {
                WriteLine("Формат записи не верный");
                goto case 1;
            }
            try{wb.Worksheets.Add($"{GroopName: ####-##-##}");}catch(Aspose.Cells.CellsException){WriteLine("Группа уже существует");goto Menu;}
            WriteLine("Группа создана");
            goto Menu;
        case 2:
            Write("Введите полное название группы: ");
            GroopName = ReadLine();
            GroopName.ToUpper();
            if(!ChekerGroop(GroopName))
            {
                WriteLine("Формат записи не верный");
                goto case 2;
            }
                try{GroopName = wb.Worksheets[GroopName].Name;}catch(Exception){WriteLine("Данной группы не существует"); goto case 2;}
                wb.Worksheets.RemoveAt(GroopName);
                WriteLine("Группа удаленна");
                cheek++;
                goto Menu;
        case 3:
            try{
                WriteLine("{0,-2} | {1}","ID", "Группы / Потоки");
                for (int i = 0; i < countSheet; i++)
                {   
                    idgroop++;
                    if(wb.Worksheets[i].Name.StartsWith("Evaluation")) // теперь легче считывать группы / потоки
                    {
                        idgroop--;
                        continue;
                    }
                    WriteLine($"{idgroop,-2}"+" | "+wb.Worksheets[i].Name);
                }
            }catch(System.ArgumentOutOfRangeException){WriteLine("Ошибка: выход индекса за предела");}
            catch(System.TypeInitializationException){WriteLine("Непонятная ошибка");}
            goto Menu;
        case 4:
            Write("Введите название группы: ");
            GroopName = ReadLine();
            GroopName.ToUpper();
            if(GroopName.Length != 10)
            {
                WriteLine("Формат записи не верный");
                goto case 4;
            }
            if(!ChekerGroop(GroopName))
            {
                WriteLine("Формат записи не верный");
                goto case 4;
            }
            try{var exist = wb.Worksheets[GroopName].Name;}catch(System.NullReferenceException){WriteLine("Группы / потока не существует"); goto case 4;}

            Cell cell;
            String NameStudent;
            for (int j = 0; j < 30; j++)
            {
                if (wb.Worksheets[GroopName].Cells.MaxDataRow-1 != 0)
                {
                    j = wb.Worksheets[GroopName].Cells.MaxDataRow+1;
                    if(j > 29){
                        WriteLine("В группе уже 30 студентов");
                        break;
                    }
                }
            student:
                Write($"{j+1} - Введите имя студента: ");
                NameStudent = ReadLine();
                if (CheckName(NameStudent) == false)
                {
                    WriteLine("Имя должны содержать только буквы!");
                    goto student;
                }
                cell = wb.Worksheets[GroopName].Cells[j,0];
                cell.PutValue(NameStudent);
                if (NameStudent?.Length <= 0)
                {
                    WriteLine("Запись закончена");
                    wb.Worksheets[GroopName].Cells.DeleteBlankRows();
                    goto Menu;
                }
            }
            AutoSortName(GroopName,30);
            wb.Worksheets[GroopName].Cells.DeleteBlankRows();
            goto Menu;
        case 5:
            Write("Введите название группы / потока: ");
            GroopName = ReadLine();
            GroopName.ToUpper();

            try{var exist = wb.Worksheets[GroopName].Name;}catch(System.NullReferenceException){WriteLine("Группы / потока не существует"); goto case 5;}

            if (wb.Worksheets[GroopName].Cells.MaxDataRow > 30)
            {
                WriteLine("|{0,-3} | {1,-35} | {2,-10}|","ID","ФИО студента","Группа");
            }
            else
            {
                WriteLine("|{0,-2} | {1,-35}|","ID","ФИО студента");
            }

            for (int j = 0; j <= wb.Worksheets[GroopName].Cells.MaxDataRow; j++)
            {  
                if(wb.Worksheets[GroopName].Cells[j,0].StringValue.Length == 0){
                    goto Menu;
                } 
                if(wb.Worksheets[GroopName].Cells.MaxDataRow > 30){
                    var cell1 = wb.Worksheets[GroopName].Cells[j,0];
                    var cell2 = wb.Worksheets[GroopName].Cells[j,1];
                    WriteLine($"|{(j+1),-3} | {cell1.StringValue,-35} | {cell2.StringValue,-10}|");
                }else{
                    cell = wb.Worksheets[GroopName].Cells[j,0];
                    string frame = string.Concat(Enumerable.Repeat("-",cell.StringValue.Length));
                    WriteLine($"|{(j+1),-2}"+" | "+ $"{cell.StringValue,-35}" + "|");
                }
            }
            goto Menu;
        case 6:
            Write("Введите название группы: ");
            GroopName = ReadLine();
            if(GroopName.Length < 10){WriteLine("Неверна задана группа"); goto case 6;}
            try{var exist = wb.Worksheets[GroopName].Name;}catch(System.NullReferenceException){WriteLine("Группы не существует"); goto case 6;}
            Write("Введите номер студента из его списка групп: ");
            int NameStudentNum = Convert.ToInt32(ReadLine());
            if(NameStudentNum > 30){
                WriteLine("В группе максимум 30 студентов");
                goto case 6;
            }
            wb.Worksheets[GroopName].Cells.DeleteRow(NameStudentNum-1);
            wb.Worksheets[GroopName].Cells.DeleteBlankRows();
            goto Menu;
        case 7:
            Write("Введите название потока без года формата ББББ: ");
            String ThreadName = ReadLine();
            ThreadName.ToUpper();
            if(ThreadName.Length != 4)
            {
                WriteLine("Неверный формат записи названия потока");
                goto case 7;
            }
            for (int i = 0; i < 4; i++)
            {
                if(!char.IsLetter(ThreadName[i]))
                {
                    WriteLine("Неверный формат записи названия потока"); 
                    goto case 7;
                }
            }
            Write("Введите год потока из последних двух цифр: ");
            String EarsThread = ReadLine();
            if(EarsThread.Length != 2)
            {
                WriteLine("Неверный формат записи года потока");
                goto case 7;
            }
            for (int i = 0; i < 2; i++)
            {
                if(!char.IsNumber(EarsThread[i]))
                {
                    WriteLine("Неверный формат записи года потока"); 
                    goto case 7;
                }
            }
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
            goto Menu;
        case 8:
            Write("Введите название группы / потока: ");
            GroopName = ReadLine();
            GroopName.ToUpper();
            try{var exist = wb.Worksheets[GroopName].Name;}catch(System.NullReferenceException){WriteLine("Группы не существует"); goto case 8;}
            int countrow = wb.Worksheets[GroopName].Cells.MaxDataRow;
            Write("Вы желаете отсортировать поток по Алфваиту ФИО или по группам (1 - по ФИО, 2 - по группам, другую любую цифру - оставить без изменений): ");
            ret1:
            try{int s = Int32.Parse(ReadLine());
            if(s == 1){
                AutoSortName(GroopName,countrow);
            }else if(s == 2){
                AutoSortGrooup(GroopName,countrow);
            }else{
                WriteLine("ну лан");
            }}
            catch(Exception){
                WriteLine("Ошибка ввода");
                goto ret1;
                }
            goto Menu;
        case 9:
            try{wb.Save(NameFileXLS);}catch(Exception){WriteLine("Error save");break;}
            WriteLine("Сохранение успешно!");
            goto Menu;
        default:
            break;
    }

} while (n != 0);

bool ChekerGroop(string groopName)
{
    if(groopName.Length != 10){return false;}
    for (int i = 0; i < 4; i++)
    {
        if(!char.IsLetter(groopName[i]))
        {
            return false; 
            //break;
        }
    }
    if(groopName[4] != '-' || groopName[7] != '-'){return false;}
    if(char.IsNumber(groopName[5]) 
        && (char.IsNumber(groopName[6]) && groopName[6] != '0') 
        && char.IsNumber(groopName[8]) 
        && (char.IsNumber(groopName[9])))
    {return true;}
    else{return false;}
}

void AutoSortName(string GroopName,int MaxRow){
    DataSorter sorter = wb.DataSorter;
    sorter.Order1 = SortOrder.Ascending;
    sorter.Key1 = 0;
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
    CellArea ca = new CellArea();
    ca.StartRow = 0;
    ca.StartColumn = 0;
    ca.EndRow = MaxRow;
    ca.EndColumn = 1;
    sorter.Sort(wb.Worksheets[GroopName].Cells, ca);
    wb.Worksheets[GroopName].Cells.DeleteBlankRows();
}
bool CheckName(String StudentName){
    StudentName.Split(" ");
    foreach (char word in StudentName)
    {
        if(char.IsNumber(word)){
            return false;
        }
    }
    return true;
}