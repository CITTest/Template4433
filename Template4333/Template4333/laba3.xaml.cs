using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Linq;
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
namespace Template4333
{
    /// <summary>
    /// Логика взаимодействия для laba3.xaml
    /// </summary>
    public partial class laba3 : Window
    {
        public laba3()
        {
            InitializeComponent();
        }

        private void BtnImp_Click(object sender, RoutedEventArgs e)
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
            int _rows = (int)lastCell.Row-11;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (laba33Entities4 laba33 = new laba33Entities4())
            {
                for (int i = 1; i < _rows; i++)
                {
                    laba33.laba3isr.Add(new laba3isr()
                    {
                        kod_zakaza = list[i, 1],
                        Data_sozdania =list[i, 2],
                        Vremya_zakaza = list[i, 3],
                        kod_klienta = list[i, 4],
                        uslugi = list[i, 5],
                        statuz = list[i, 6],
                        data_zakritia = list[i, 7],
                        vremya_prokata = list[i, 8]
                    });
                }
                laba33.SaveChanges();
            }
        }

        private void BtnExp_Click(object sender, RoutedEventArgs e)
        {
            List<laba3isr> alldannie;
            List<string> strings;
            using (laba33Entities4 laba33 = new laba33Entities4())
            {
                alldannie =
                laba33.laba3isr.ToList().OrderBy(s =>
                s.vremya_prokata).ToList();
                strings =
                laba33.laba3isr.ToList().Select(laba3isr => laba3isr.vremya_prokata.ToString()).Distinct().ToList();
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
                foreach (var dann in alldannie) 
                {
                    if(dann.vremya_prokata == strings[i])
                    {
                        worksheet.Cells[1][startRowIndex] = dann.id.ToString();
                        worksheet.Cells[2][startRowIndex] = dann.kod_zakaza;
                        worksheet.Cells[3][startRowIndex] = dann.Data_sozdania;
                        worksheet.Cells[4][startRowIndex] = dann.kod_klienta;
                        worksheet.Cells[5][startRowIndex] = dann.uslugi;
                        startRowIndex++;
                    }
                }
            }
            app.Visible = true;
            BtnExp.Background = new SolidColorBrush(Colors.Black);
        }
    }
}