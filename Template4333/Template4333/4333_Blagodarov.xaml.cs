using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

namespace Template4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Blagodarov.xaml
    /// </summary>
    public partial class _4333_Blagodarov : Window
    {
        public _4333_Blagodarov()
        {
            InitializeComponent();
        }

        int NumberOfDateRow;
        DataTable table = new DataTable();

        public class StringManipulation
        {
            public string RemoveCharactersAfterLastSpaceRight(string input)
            {
                int lastSpaceIndex = input.LastIndexOf(' '); // Находим позицию последнего пробела

                if (lastSpaceIndex != -1)
                {
                    return input.Substring(0, lastSpaceIndex + 1); // Обрезаем строку до последнего пробела включительно
                }

                return input; // Если нет пробела, возвращаем исходную строку
            }
            public string RemoveCharactersBeforeFirstSpaceLeft(string input)
            {
                int firstSpaceIndex = input.IndexOf(' '); // Находим позицию первого пробела

                if (firstSpaceIndex != -1)
                {
                    return input.Substring(firstSpaceIndex + 1); // Обрезаем строку после первого пробела
                }

                return input; // Если нет пробела, возвращаем исходную строку
            }
        }
        public void ImportDataFromExcel(string filePath)
        {
            if (filePath == null)
            {
                return;
            }

            table.Columns.Clear();
            table.Rows.Clear();

            int columnNumber = 0;

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
            Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            Excel.Range excelRange = excelWorksheet.UsedRange;
            StringManipulation manipulation = new StringManipulation();

            // Создаем колонки
            for(int i = 1; i < 50; i++)
            {
                string column = excelRange.Cells[1, i].Value != null ? excelRange.Cells[1, i].Value.ToString() : null;
                if (column != null)
                {
                    columnNumber++;
                    table.Columns.Add(column,typeof(string));
                } else { break; }
            }

            int rows = excelRange.Rows.Count;

            for (int i = 2; i < rows; i++)
            {
                DataRow newRow = table.NewRow();
                for (int j = 1; j <= columnNumber; j++)
                {
                    string row = excelRange.Cells[i, j].Value != null ? excelRange.Cells[i, j].Value.ToString() : null;
                    if (row != null && row.ToString() != "")
                    {
                        //Преобразовываем все возможные ячейки данных в верный формат DateTime
                        if (DateTime.TryParse(row, out DateTime dateValue)) // Пытаемся преобразовать строку в DateTime
                        {
                            newRow[excelRange.Cells[1, j].Value.ToString()] = dateValue; // Если удалось конвертировать, записываем новое значение обратно
                            NumberOfDateRow = j;
                        } else
                        {
                            newRow[excelRange.Cells[1, j].Value.ToString()] = row; // Если не удалось, то сохраняем прошлое значение обратно
                        }
                    } else { break; }
                }
                table.Rows.Add(newRow);
            }

            
            /*foreach (DataRow row in table.Rows)
            {
                string dateString = row[NumberOfDateRow].ToString(); // Получаем значение ячейки столбца как строку
                if (DateTime.TryParse(dateString, out DateTime dateValue)) // Пытаемся преобразовать строку в DateTime
                {
                    row[NumberOfDateRow] = dateValue; 
                }
            }*/

            Data.ItemsSource = table.DefaultView;

            /*            // Перебор данных из Excel файла
                        for (int i = 2; i <= rows; i++)
                        {
                            int id = excelRange.Cells[i, 1].Value != null ? int.Parse(excelRange.Cells[i, 1].Value.ToString()) : 0;
                            string orderCode = excelRange.Cells[i, 2].Value != null ? excelRange.Cells[i, 2].Value.ToString() : "";
                            string customerCode = excelRange.Cells[i, 5].Value != null ? excelRange.Cells[i, 5].Value.ToString() : "";
                            string services = excelRange.Cells[i, 6].Value != null ? excelRange.Cells[i, 6].Value.ToString() : "";
                            string date = excelRange.Cells[i, 3].Value != null ? excelRange.Cells[i, 3].Value.ToString() : "01.01.0001";
                            string time = excelRange.Cells[i, 4].Value != null ? excelRange.Cells[i, 4].Value.ToString() : "0,0000";
                            object timeparse = DateTime.FromOADate(Convert.ToDouble(time));
                            string datetime = manipulation.RemoveCharactersAfterLastSpaceRight(date) + " " + manipulation.RemoveCharactersBeforeFirstSpaceLeft(timeparse.ToString());
                            DateTime createdDate = DateTime.Parse(datetime);


                            newRow["Id"] = id;
                            newRow["Код заказа"] = orderCode;
                            newRow["Код клиента"] = customerCode;
                            newRow["Услуги"] = services;
                            newRow["Дата|Время"] = createdDate;

                            table.Rows.Add(newRow);

                            Console.WriteLine($"Saving to database - Id: {id}, Order Code: {orderCode}, Customer Code: {customerCode}, Services: {services}, Created Date: {createdDate}");
                        }*/

            excelWorkbook.Close();
            excelApp.Quit();
        }
        public void ExportDataFromExcel(DataTable table, SaveFileDialog saveFileDialog)
        {
            if (table == null)
            {
                return;
            }

            DataSet dataSet = new DataSet();
            dataSet.Clear();

            int tablesNumber = 0;

            // Цикл для хранения уникальных дат в исходной таблице
            HashSet<DateTime> uniqueDates = new HashSet<DateTime>();

            // Получение уникальных дат
            foreach (DataRow row in table.Rows)
            {
                if (row[NumberOfDateRow - 1] != null && row[NumberOfDateRow - 1].ToString() != "")
                {
                    DateTime date = Convert.ToDateTime(row[NumberOfDateRow - 1]);
                    uniqueDates.Add(date);
                }
                else { continue; }
            }

            // Создание новых DataTable для каждой уникальной даты
            foreach (DateTime date in uniqueDates)
            {
                DataTable newTable = new DataTable();
                newTable.TableName = "DataForDate_" + date.ToString("yyyyMMdd");

                // Копирование структуры исходной таблицы в новую
                foreach (DataColumn col in table.Columns)
                {
                    newTable.Columns.Add(col.ColumnName, col.DataType);
                }

                // Добавление строк только с искомой датой
                foreach (DataRow row in table.Rows)
                {
                    if (row[NumberOfDateRow - 1] != null && row[NumberOfDateRow - 1].ToString() != "")
                    {
                        if (Convert.ToDateTime(row[NumberOfDateRow - 1]).Date == date.Date)
                        {
                            newTable.ImportRow(row);
                        }
                    } else { continue; }
                }

                // Добавление новой таблицы в DataSet
                tablesNumber++;
                dataSet.Tables.Add(newTable);
            }

            if (saveFileDialog.ShowDialog() == true)
            {
                // Создание нового Excel пакета
                //Excel.Workbook wb = new Excel.Workbook();
                Excel.Application excelApp = new Excel.Application();

                // Отключаем отображение Excel во время создания
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                // Создаем новую книгу Excel
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

                foreach (DataTable dataTable in dataSet.Tables)
                {
                    // Добавление нового листа с выбранным названием и шириной колонок
                    excelWorksheet = excelWorkbook.Sheets.Add();
                    excelWorksheet.Name = dataTable.TableName;
                    excelWorksheet.Columns.ColumnWidth = 20;

                    // Запись заголовков (названий столбцов DataTable) в первую строку
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        excelWorksheet.Cells[1, i].Value = dataTable.Columns[i - 1].ColumnName;
                    }

                    // Запись данных строк DataTable в Excel
                    for (int r = 0; r < dataTable.Rows.Count; r++)
                    {
                        for (int c = 0; c < dataTable.Columns.Count; c++)
                        {
                            excelWorksheet.Cells[r + 2, c + 1].Value = dataTable.Rows[r][c];
                        }
                    }
                }

                // Сохранение Excel файла по выбранному пути
                FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                excelWorkbook.SaveAs(excelFile.FullName);

                MessageBox.Show("Файл был успешно сохранен", "Successfully",MessageBoxButton.OK);

                // Закрываем и освобождаем ресурсы
                excelWorkbook.Close();
                excelApp.Quit();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ImportDataFromExcel("C:\\Users\\MSII\\OneDrive\\Рабочий стол\\КИТ\\Инструментальные средства разработки программного обеспечения\\Лабораторная работа №3\\2.xlsx");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            // Выбор места сохранения и названия Excel файла
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDialog.FileName = "ExcelFileWithCategories.xlsx";

            ExportDataFromExcel(table, saveFileDialog);

            this.Close();
        }
    }
}
