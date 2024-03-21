using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;
using System.Security.Cryptography;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Markup;



namespace Template4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_KhokhlovAlexey.xaml
    /// </summary>
    public partial class _4333_KhokhlovAlexey : Window
    {
        public _4333_KhokhlovAlexey()
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
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (ISRPO3Entities usersEntities = new ISRPO3Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.xls.Add(new xls()
                    {
                        Код_клиента = list[i, 0],
                        Должность = list[i, 1],
                        ФИО = list[i, 2],
                        Логин = list[i, 3],
                        Пароль = list[i, 4],
                        Последний_вход = list[i, 5],
                        Тип_входа = list[i, 6]
                    });
                }
                usersEntities.SaveChanges();
            }

        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<xls> alldata;
            List<string> strings;
            using (ISRPO3Entities laba3 = new ISRPO3Entities())
            {
                alldata =
                laba3.xls.ToList().OrderBy(s => s.Должность).ToList();
                strings = laba3.xls.ToList().Select(xls => xls.Должность.ToString()).Distinct().ToList();
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
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "Логин";
                startRowIndex++;
                foreach (var data in alldata)
                {
                    if (data.Должность == strings[i])
                    {
                        worksheet.Cells[1][startRowIndex] = data.Код_клиента;
                        worksheet.Cells[2][startRowIndex] = data.ФИО;
                        worksheet.Cells[3][startRowIndex] = data.Логин;
                        startRowIndex++;
                    }
                }
            }
            app.Visible = true;
            BnExport.Background = new SolidColorBrush(Colors.Black);
        }

        private async void BnImportJS_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON files (*.json)|*.json";

            if (openFileDialog.ShowDialog() == true)
            {
                string jsonFilePath = openFileDialog.FileName;

                List<xls1> alldata;

                using (FileStream fs = new FileStream(jsonFilePath, FileMode.Open))
                {
                    alldata = await JsonSerializer.DeserializeAsync<List<xls1>>(fs);
                }

                using (ISRPO3Entities2 usersEntities = new ISRPO3Entities2())
                {
                    foreach (var alldat in alldata)
                    {
                        xls1 mytable = new xls1
                        {
                            CodeStaff = alldat.CodeStaff,
                            Position = alldat.Position,
                            FullName = alldat.FullName,
                            Lg = alldat.Lg,
                            Pass = alldat.Pass,
                            LastEnter = alldat.LastEnter,
                            TypeEnter = alldat.TypeEnter,
                        };
                        usersEntities.xls1.Add(mytable);
                    }
                    usersEntities.SaveChanges();
                }


            }
        }


        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<xls1> alldata;

            using (ISRPO3Entities2 usersEntities = new ISRPO3Entities2())
            {
                alldata = usersEntities.xls1.ToList().OrderBy(s => s.Position).ToList();

            }
            foreach (var group in alldata.GroupBy(o => o.Position))
            {
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();

                Word.Paragraph headerParagraph = document.Paragraphs.Add();
                Word.Range headerRange = headerParagraph.Range;
                headerRange.Text = $"Должность: {group.Key} (Количество сотрудников: { group.Count()})";
                headerParagraph.set_Style("Заголовок 1");
                headerRange.InsertParagraphAfter();

                Word.Table xlsTable = document.Tables.Add(headerRange, group.Count() + 2, 3);
                xlsTable.Borders.InsideLineStyle = xlsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                xlsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                xlsTable.Cell(1, 1).Range.Text = "Код сотрудника";
                xlsTable.Cell(1, 2).Range.Text = "ФИО";
                xlsTable.Cell(1, 3).Range.Text = "Логин";
                //xlsTable.Cell(1, 4).Range.Text = "Количество сотрудников";

                int i = 1;
                foreach (var order in group)
                {
                    i++;
                    xlsTable.Cell(i, 1).Range.Text = order.CodeStaff;
                    xlsTable.Cell(i, 2).Range.Text = order.FullName;
                    xlsTable.Cell(i, 3).Range.Text = order.Lg;
                    //xlsTable.Cell(i, 4).Range.Text = group.Count().ToString();
                }
                xlsTable.Cell(group.Count() + 2, 1).Range.Text = $"Всего сотрудников: {group.Count()}";
                xlsTable.Cell(group.Count() + 2, 1).Merge(xlsTable.Cell(group.Count() + 2, 3));


                string fileName = $"C:/Users/Alexey/Desktop/Учёба/ИСРПО/outputFileWord_{group.Key}_{group.Count()}.docx";
                document.SaveAs2(fileName);

                app.Visible = true;

            }
        }
    }
}

