using System.Windows;
using System.IO;
using OfficeOpenXml;
using System.Data.SqlClient;
using Microsoft.Win32;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Data.Entity;
using System.Data;
using System.Windows.Documents;
using System.Text.Json;
using System.Reflection;
using System.Windows.Controls;

namespace Template4333
{
    public partial class _4333_Demidov : Window
    {
        public _4333_Demidov()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        public DataTable Table12;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\USERS\YURIC\DOWNLOADS\TEMPLATE4333\TEMPLATE4333\TEMPLATE4333\DATABASE.MDF;Integrated Security=True");
            connection.Open();
            string[,] list;
            int _columns = 4;
            int _rows = 10;
            list = new string[_rows, _columns];
            SqlCommand cmd = new SqlCommand("Select id, NSP, Login ,Statement From Table12", connection);
            SqlDataReader dr = cmd.ExecuteReader();
            int a = 0;
            while (dr.Read())
            {
                list[a, 0] = dr[0].ToString();
                list[a, 1] = dr[1].ToString();
                list[a, 2] = dr[2].ToString();
                list[a, 3] = dr[3].ToString();
                a++;
            }
            dr.Close();

            List<string> positions = new List<string>();

            for (int i = 0; i < list.GetLength(0); i++)
            {
                if (!positions.Contains(list[i, 0]))
                {
                    positions.Add(list[i, 0]);
                }
            }

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add();
            Excel.Worksheet currentSheet = null;
            int k = 1;
            int rowToStart = 1;
            for (int i = 0; i < list.GetLength(0); i++)
            {
                string position = list[i, 3];
                Excel.Worksheet positionSheet = null;

                foreach (Excel.Worksheet sheet in xlWorkBook.Sheets)
                {
                    if (sheet.Name == position)
                    {
                        positionSheet = sheet;
                        for (int c = 1; c < list.GetLength(0); c++)
                        {
                            string currentSheetFirstCell = (string)(positionSheet.Cells[c, 1] as Excel.Range).Value;
                            if (currentSheetFirstCell == null || currentSheetFirstCell == "")
                            {
                                rowToStart = c;
                                break;
                            }
                        }
                        break;
                    }
                }

                if (positionSheet == null)
                {
                    positionSheet = (Excel.Worksheet)xlWorkBook.Sheets.Add();
                    positionSheet.Name = position;
                    rowToStart = 1;
                }

                currentSheet = positionSheet;
                int rowToEnd = 1;
                int counter = 1;
                for (int j = 0; j < list.GetLength(1); j++)
                {
                    currentSheet.Cells[rowToStart, rowToEnd] = list[i, j];
                    rowToEnd++;
                    if (counter % 4 == 0)
                    {
                        rowToStart++;
                    }
                    counter++;
                    
                }
            }

            xlWorkBook.SaveAs("C:\\Users\\yuric\\Downloads\\BebraISRPO3.xlsx");
            xlWorkBook.Close();
            xlApp.Quit();
            GC.Collect();
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "Файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };

            // Если файл выбран
            if (ofd.ShowDialog() == true)
            {
                string[,] list;
                Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                int _columns = (int)lastCell.Column;
                int _rows = (int)lastCell.Row;
                list = new string[_rows, _columns];

                for (int i = 0; i < _rows; i++)
                {
                    for (int j = 0; j < _columns; j++)
                    {
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }

                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
                using (SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;Initial Catalog=Database;Integrated Security=True"))
                {
                    connection.Open();
                    SqlCommand sql = new SqlCommand("Delete From Table12", connection);
                    SqlDataReader dr = sql.ExecuteReader();
                    dr.Close();
                    for (int h = 1; h < _rows; h++)
                    {
                        DateTime parsedDate = DateTime.ParseExact(list[h,5], "dd:MM:yyyy HH:mm:ss", null);

                        string formattedDate = parsedDate.ToString("yyyy-MM-dd HH:mm:ss");
                        SqlCommand cmd = new SqlCommand("INSERT INTO Table12 VALUES('" + list[h, 0].ToString() + "',N'" + list[h, 1].ToString() + "',N'" + list[h, 2].ToString() + "'," +
                                "'" + list[h, 3].ToString() + "','" + list[h, 4].ToString() + "','" + formattedDate + "',N'" + list[h, 6].ToString() + "');", connection);
                        dr = cmd.ExecuteReader();
                        dr.Close ();
                    }
                }
            }
        }
        private void Button_WordClick_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "Файл Json (Spisok.json)|*.json",
                Title = "Выберите файл базы данных"
            };
            if (ofd.ShowDialog() == true)
            {
                string jsonData = File.ReadAllText(ofd.FileName);

                Table12 = JsonConvert.DeserializeObject<DataTable>(jsonData);
                using (SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;Initial Catalog=Database;Integrated Security=True"))
                {
                    connection.Open();
                    SqlCommand sql = new SqlCommand("Delete From Table12", connection);
                    SqlDataReader dr = sql.ExecuteReader();
                    dr.Close();
                    for (int h = 0; h < Table12.Rows.Count; h++)
                    {
                        DataRow row1 = Table12.Rows[h];
                        DateTime parsedDate = DateTime.ParseExact(row1[6].ToString(), "dd:MM:yyyy HH:mm:ss", null);

                        string formattedDate = parsedDate.ToString("yyyy-MM-dd HH:mm:ss");
                        SqlCommand cmd = new SqlCommand("INSERT INTO Table12 VALUES('" + row1[1].ToString() + "',N'" + row1[2].ToString() + "',N'" + row1[3].ToString() + "'," +
                                "'" + row1[4].ToString() + "','" + row1[5].ToString() + "','" + formattedDate + "',N'" + row1[7].ToString() + "');", connection);
                        dr = cmd.ExecuteReader();
                        dr.Close();
                    }
                }
            }
        }
        private void Button_WordClick_2(object sender, RoutedEventArgs e)
        {
            SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;Initial Catalog=Database;Integrated Security=True");
            connection.Open();
            string[,] list;
            int _columns = 4;
            int _rows = 10;
            list = new string[_rows, _columns];
            SqlCommand cmd = new SqlCommand("Select id, SNP, Login ,Statement From Table12", connection);
            SqlDataReader dr = cmd.ExecuteReader();
            int a = 0;
            while (dr.Read())
            {
                list[a, 0] = dr[0].ToString();
                list[a, 1] = dr[1].ToString();
                list[a, 2] = dr[2].ToString();
                list[a, 3] = dr[3].ToString();
                a++;
            }
            dr.Close();

            List<string> positions = new List<string>();

            for (int i = 0; i < list.GetLength(0); i++)
            {
                if (!positions.Contains(list[i, 3]))
                {
                    positions.Add(list[i, 3]);
                }
            }
            Dictionary<string, int> employeeCountByPosition = new Dictionary<string, int>();

            foreach (string position in positions)
            {
                int count = 0;
                for (int i = 0; i < list.GetLength(0); i++)
                {
                    if (list[i, 3] == position)
                    {
                        count++;
                    }
                }
                employeeCountByPosition.Add(position, count);
            }

            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            wordApp.Visible =  false;
            foreach (var position in positions)
            {
                Word.Range range = wordApp.Selection.Range;

                // Добавляем новую страницу
                wordApp.Selection.InsertNewPage();
                //wordDoc.Range().InsertAfter($"Должность: {position}");

                //wordDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                // Таблица для вывода информации
                Word.Table wordTable = wordDoc.Tables.Add(range, 1, 4);
                wordTable.Borders.Enable = 1;

                // Заголовки столбцов
                wordTable.Cell(1, 1).Range.Text = "ID";
                wordTable.Cell(1, 2).Range.Text = "ФИО";
                wordTable.Cell(1, 3).Range.Text = "Логин";
                wordTable.Cell(1, 4).Range.Text = "Должность";

                int rowIndex = 2;
                for (int i = 0; i < list.GetLength(0); i++)
                {
                    if (list[i, 3] == position)
                    {
                        wordTable.Rows.Add();
                        wordTable.Cell(rowIndex, 1).Range.Text = list[i, 0];
                        wordTable.Cell(rowIndex, 2).Range.Text = list[i, 1];
                        wordTable.Cell(rowIndex, 3).Range.Text = list[i, 2];
                        wordTable.Cell(rowIndex, 4).Range.Text = list[i, 3];
                        rowIndex++;
                    }

                }
                int employeesCount = wordTable.Rows.Count - 1;
                wordTable.Rows.Add();
                wordTable.Cell(rowIndex, 1).Range.Text = "Должность: " + position;
                wordTable.Cell(rowIndex, 2).Range.Text = "Количество сотрудников: " + employeesCount;
                wordTable.Cell(rowIndex, 1).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                wordTable.Cell(rowIndex, 1).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Cell(rowIndex, 2).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                wordTable.Cell(rowIndex, 2).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Cell(rowIndex, 3).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                wordTable.Cell(rowIndex, 3).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                wordTable.Cell(rowIndex, 4).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                wordTable.Cell(rowIndex, 4).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            }
            wordDoc.SaveAs(@"C:\Users\yuric\Downloads\BebraISRPO33.docx");
            wordDoc.Close();
            wordApp.Application.Quit();
            wordApp.Quit();
        }
    }
}
