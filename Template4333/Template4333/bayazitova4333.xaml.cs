using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
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
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;


namespace Template4333
{
    /// <summary>
    /// Логика взаимодействия для bayazitova4333.xaml
    /// </summary>
    public partial class bayazitova4333 : Window
    {
        public bayazitova4333()
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

            using (newdbEntities usersEntities = new newdbEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.Infoes.Add(new Info()
                    {

                        Name = list[i, 1],
                        Type_of_service = list[i, 2],
                        Id_service = list[i, 3],
                        Cost_rub = Convert.ToInt32(list[i, 4])

                    });
                }
                usersEntities.SaveChanges();
            }

        }

        private void BnExp_Click(object sender, RoutedEventArgs e)
        
            {
                List<Info> alldannie;
            List<Info> dannie1;
            List<Info> dannie2;
            List<Info> dannie3;
            using (newdbEntities bay33 = new newdbEntities())
                {
                    alldannie =
                    bay33.Infoes.ToList().OrderBy(s =>
                    s.Cost_rub).ToList();
                    dannie1 =
                    bay33.Infoes.OrderBy(s =>
                    s.Cost_rub).Where(s =>s.Cost_rub<=250 && s.Cost_rub>=0).ToList();
                    dannie2 =
                    bay33.Infoes.OrderBy(s =>
                    s.Cost_rub).Where(s => s.Cost_rub <= 800 && s.Cost_rub >= 250).ToList();
                    dannie3 =
                    bay33.Infoes.OrderBy(s =>
                    s.Cost_rub).Where(s => s.Cost_rub > 800).ToList();

            }
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = 3;
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[1];
                worksheet.Name = 1.ToString();
                worksheet.Cells[1][startRowIndex] = "ID";
                worksheet.Cells[2][startRowIndex] = "Название услуги";
                worksheet.Cells[3][startRowIndex] = "Вид улслуги";
                worksheet.Cells[4][startRowIndex] = "Стоимость";
                startRowIndex++;
                foreach (var dann in dannie1)
                {
                    worksheet.Cells[1][startRowIndex] = dann.ID.ToString();
                    worksheet.Cells[2][startRowIndex] = dann.Name;
                    worksheet.Cells[3][startRowIndex] = dann.Type_of_service;
                    worksheet.Cells[4][startRowIndex] = dann.Cost_rub.ToString();
                    startRowIndex++;
            }
            startRowIndex = 1;
            Excel.Worksheet worksheet1 = app.Worksheets.Item[2];
            worksheet.Name = 1.ToString();
            worksheet1.Cells[1][startRowIndex] = "ID";
            worksheet1.Cells[2][startRowIndex] = "Название услуги";
            worksheet1.Cells[3][startRowIndex] = "Вид улслуги";
            worksheet1.Cells[4][startRowIndex] = "Стоимость";
            startRowIndex++;
            foreach (var dann in dannie2)
            {
                worksheet1.Cells[1][startRowIndex] = dann.ID.ToString();
                worksheet1.Cells[2][startRowIndex] = dann.Name;
                worksheet1.Cells[3][startRowIndex] = dann.Type_of_service;
                worksheet1.Cells[4][startRowIndex] = dann.Cost_rub.ToString();
                startRowIndex++;
            }
            startRowIndex = 1;
            Excel.Worksheet worksheet2 = app.Worksheets.Item[3];
            worksheet.Name = 1.ToString();
            worksheet2.Cells[1][startRowIndex] = "ID";
            worksheet2.Cells[2][startRowIndex] = "Название услуги";
            worksheet2.Cells[3][startRowIndex] = "Вид улслуги";
            worksheet2.Cells[4][startRowIndex] = "Стоимость";
            startRowIndex++;
            foreach (var dann in dannie3)
            {
                worksheet2.Cells[1][startRowIndex] = dann.ID.ToString();
                worksheet2.Cells[2][startRowIndex] = dann.Name;
                worksheet2.Cells[3][startRowIndex] = dann.Type_of_service;
                worksheet2.Cells[4][startRowIndex] = dann.Cost_rub.ToString();
                startRowIndex++;
            }

            app.Visible = true;

        }

        private async void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "JSON files (*.json)|*.json";

                if (openFileDialog.ShowDialog() == true)
                {
                    string jsonFilePath = openFileDialog.FileName;

                    List<Info> ordersData;

                    using (FileStream fs = new FileStream(jsonFilePath, FileMode.Open))
                    {
                        ordersData = await JsonSerializer.DeserializeAsync<List<Info>>(fs);
                    }

                    using (newdbEntities usersEntities = new newdbEntities())
                    {
                        foreach (var orderData in ordersData)
                        {
                            Info newOrder = new Info
                            {
                               
                                Name = orderData.Name,
                                Type_of_service = orderData.Type_of_service,
                                Id_service = orderData.Id_service,
                                Cost_rub = orderData.Cost_rub,
                               
                            };
                            usersEntities.Infoes.Add(newOrder);
                        }
                        usersEntities.SaveChanges();
                    }

                    MessageBox.Show("Данные успешно импортированы из JSON файла в таблицу БД", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex) { MessageBox.Show("Произошла ошибка при добавлении данных: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private void BtnExp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<Info> alldannie;
                List<Info> dannie1;
                List<Info> dannie2;
                List<Info> dannie3;

                using (newdbEntities bay33 = new newdbEntities())
                {
                    alldannie = bay33.Infoes.ToList().OrderBy(s => s.Cost_rub).ToList();
                    dannie1 = bay33.Infoes.OrderBy(s => s.Cost_rub).Where(s => s.Cost_rub <= 250 && s.Cost_rub >= 0).ToList();
                    dannie2 = bay33.Infoes.OrderBy(s => s.Cost_rub).Where(s => s.Cost_rub <= 800 && s.Cost_rub > 250).ToList();
                    dannie3 = bay33.Infoes.OrderBy(s => s.Cost_rub).Where(s => s.Cost_rub > 800).ToList();
                }

                List<List<Info>> allGroups = new List<List<Info>>() { dannie1, dannie2, dannie3 };

                for (int i = 0; i < allGroups.Count; i++)
                {
                    var app = new Word.Application();
                    Word.Document document = app.Documents.Add();

                    Word.Paragraph headerParagraph = document.Paragraphs.Add();
                    Word.Range headerRange = headerParagraph.Range;
                    headerRange.Text = $"Группа {i + 1}: Данные по стоимости";
                    headerParagraph.set_Style("Заголовок 1");
                    headerRange.InsertParagraphAfter();

                    Word.Table dataTable = document.Tables.Add(headerRange, allGroups[i].Count + 1, 4);
                    dataTable.Borders.InsideLineStyle = dataTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    dataTable.Rows[1].Range.Font.Bold = 1;
                    dataTable.Rows[1].Range.Font.Italic = 1;
                    dataTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    dataTable.Cell(1, 1).Range.Text = "ID";
                    dataTable.Cell(1, 2).Range.Text = "Название услуги";
                    dataTable.Cell(1, 3).Range.Text = "Вид услуги";
                    dataTable.Cell(1, 4).Range.Text = "Стоимость";

                    int rowIndex = 1;
                    foreach (var item in allGroups[i])
                    {
                        rowIndex++;
                        dataTable.Cell(rowIndex, 1).Range.Text = item.ID.ToString();
                        dataTable.Cell(rowIndex, 2).Range.Text = item.Name;
                        dataTable.Cell(rowIndex, 3).Range.Text = item.Type_of_service;
                        dataTable.Cell(rowIndex, 4).Range.Text = item.Cost_rub.ToString();
                    }

                    string fileName = $"C:/Users/Asus/Desktop/export/outputFileWord_Group{i + 1}.docx";
                    document.SaveAs2(fileName);

                    app.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при экспорте данных: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
    }

