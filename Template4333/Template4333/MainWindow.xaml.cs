using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
                        ID = Convert.ToInt32(list[i, 0]),
                        Code_order = list[i, 1],
                        Data = list[i, 2],
                        Time = list[i, 3],
                        Code_client = list[i, 4],
                        Uslugi = list[i, 5],
                        Status = list[i, 6],
                        Data_close = list[i, 7],
                        Time_prokata = list[i, 8]
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
 s.Time_prokata).ToList();
                strings = usersEntities.Users.ToList().Select(Users =>
Users.Time_prokata.ToString()).Distinct().ToList();
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
                    if (students.Time_prokata == strings[i])
                    {
                        worksheet.Name = strings[i];
                        worksheet.Cells[1][startRowIndex] = students.ID.ToString();
                        worksheet.Cells[2][startRowIndex] = students.Code_order;
                        worksheet.Cells[3][startRowIndex] = students.Data;
                        worksheet.Cells[4][startRowIndex] = students.Code_client;
                        worksheet.Cells[5][startRowIndex] = students.Uslugi;
                        startRowIndex++;
                    }

                }
            }

            app.Visible = true;

        }


    }
}



