using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
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
            for (int j = 0; j<_columns; j++)
                 for (int i = 0; i<_rows; i++)
                list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
             ObjWorkBook.Close(false, Type.Missing, Type.Missing);
             ObjWorkExcel.Quit();
             GC.Collect();
            using (ISRPO3Entities usersEntities = new ISRPO3Entities())
            {
                for (int i = 0; i<_rows; i++)
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
    }
}

