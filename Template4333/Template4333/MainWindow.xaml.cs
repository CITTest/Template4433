using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Template4333;
using Excel = Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;
using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System.Data.Entity.Validation;
using Newtonsoft.Json.Linq;




namespace InteropExcelApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (usersEntities usersEntities = new usersEntities())
            {
                for (int i = 1; i < 51; i++)
                {
                    usersEntities.Users.Add(new User()
                    {
                        Id = Convert.ToInt32(list[i, 0]),
                        CodeOrder = list[i, 1],
                        CreateDate = list[i, 2],
                        CreateTime = list[i, 3],
                        CodeClient = list[i, 4],
                        Services = list[i, 5],
                        Status = list[i, 6],
                        ClosedDate = list[i, 7],
                        ProkatTime = list[i, 8]
                    });
                }
                usersEntities.SaveChanges();
            }
            MessageBox.Show("импорт завершён");

        }
        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<User> allStudents;
            List<string> strings;
            using (usersEntities usersEntities = new usersEntities())
            {
                allStudents =
 usersEntities.Users.ToList().OrderBy(s =>
 s.ProkatTime).ToList();
                strings = usersEntities.Users.ToList().Select(Users =>
Users.ProkatTime.ToString()).Distinct().ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = strings.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < strings.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i +
                1];
                worksheet.Name = strings[i];
                worksheet.Cells[1][startRowIndex] = "ID";
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][startRowIndex] = "Дата создания";
                worksheet.Cells[4][startRowIndex] = "Код клиента";
                worksheet.Cells[5][startRowIndex] = "Услуги";
                startRowIndex++;
                foreach (var students in allStudents)
                {
                    if (students.ProkatTime == strings[i])
                    {
                        worksheet.Name = strings[i];
                        worksheet.Cells[1][startRowIndex] = students.Id.ToString();
                        worksheet.Cells[2][startRowIndex] = students.CodeOrder;
                        worksheet.Cells[3][startRowIndex] = students.CreateDate;
                        worksheet.Cells[4][startRowIndex] = students.CodeClient;
                        worksheet.Cells[5][startRowIndex] = students.Services;
                        startRowIndex++;
                    }

                }
            }

            app.Visible = true;

        }
        private void BnImportjson(object sender, RoutedEventArgs e)
{
    OpenFileDialog ofd = new OpenFileDialog()
    {
        DefaultExt = "*.json",
        Filter = "JSON файл (*.json)|*.json",
        Title = "Выберите файл базы данных JSON"
    };

    if (!(ofd.ShowDialog() == true))
        return;

    string jsonText = File.ReadAllText(ofd.FileName);

    // Десериализация JSON в объект
    var data = JsonConvert.DeserializeObject<List<User>>(jsonText);

    using (usersEntities usersEntities = new usersEntities())
    {
        int count = Math.Min(data.Count, 51); // Берем минимум из количества элементов в файле и 50

        for (int i = 0; i < count; i++)
        {
            var item = data[i];
            usersEntities.Users.Add(new User()
            {
                Id = item.Id,
                CodeOrder = item.CodeOrder,
                CreateDate = item.CreateDate,
                CreateTime = item.CreateTime,
                CodeClient = item.CodeClient,
                Services = item.Services,
                Status = item.Status,
                ClosedDate = item.ClosedDate,
                ProkatTime = item.ProkatTime
            });
        }

        try
        {
            usersEntities.SaveChanges();
        }
        catch (DbEntityValidationException ex)
        {
            // Выводим подробности об ошибках валидации
            foreach (var validationErrors in ex.EntityValidationErrors)
            {
                foreach (var validationError in validationErrors.ValidationErrors)
                {
                    MessageBox.Show($"Сущность {validationErrors.Entry.Entity.GetType().Name} ошибка: {validationError.ErrorMessage}");
                }
            }
        }
    }

    MessageBox.Show("Импорт завершен");
}




        private void BnExportWord(object sender, RoutedEventArgs e)
        {
            // Получаем список студентов и групп
            List<User> allStudents;
            List<string> prokatTimes;
            using (usersEntities usersEntities = new usersEntities())
            {
                allStudents = usersEntities.Users.ToList().OrderBy(s => s.ProkatTime).ToList();
                prokatTimes = usersEntities.Users.Select(user => user.ProkatTime).Distinct().ToList();
            }

            // Создаем новый документ Word
            Word.Application app = new Word.Application();
            Word.Document document = app.Documents.Add();

            foreach (string prokatTime in prokatTimes)
            {
                // Добавляем новую страницу для каждой категории
                Word.Paragraph paragraph = document.Paragraphs.Add();
                paragraph.Range.Text = $"Прокат по времени: {prokatTime}";
                paragraph.Range.InsertParagraphAfter();

                // Добавляем таблицу
                Word.Table table = document.Tables.Add(paragraph.Range, allStudents.Count + 1, 5); // 5 колонок для ID, Код заказа, Дата создания, Код клиента, Услуги

                // Добавляем заголовки колонок
                table.Cell(1, 1).Range.Text = "ID";
                table.Cell(1, 2).Range.Text = "Код заказа";
                table.Cell(1, 3).Range.Text = "Дата создания";
                table.Cell(1, 4).Range.Text = "Код клиента";
                table.Cell(1, 5).Range.Text = "Услуги";

                // Заполняем таблицу данными
                int rowIndex = 2;
                foreach (var student in allStudents)
                {
                    if (student.ProkatTime == prokatTime)
                    {
                        table.Cell(rowIndex, 1).Range.Text = student.Id.ToString();
                        table.Cell(rowIndex, 2).Range.Text = student.CodeOrder;
                        table.Cell(rowIndex, 3).Range.Text = student.CreateDate.ToString();
                        table.Cell(rowIndex, 4).Range.Text = student.CodeClient;
                        table.Cell(rowIndex, 5).Range.Text = student.Services;
                        rowIndex++;
                    }
                }
            }

            // Отображаем приложение Word
            app.Visible = true;

            // Сохраняем документ Word
            string savePath = "C:\\Users\\RFkzn\\Desktop\\ИСРПО4";
            document.SaveAs2($"{savePath}.docx");
            document.SaveAs2($"{savePath}.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }
    }
}



