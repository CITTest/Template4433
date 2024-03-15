using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4333
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _4333_Kulikova win = new _4333_Kulikova();  
            win.Show();
        }

        private void Kahraman_Click(object sender, RoutedEventArgs e)
        {
            _4333_Kahraman kahraman = new _4333_Kahraman();
            kahraman.Show();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = ".xls;*.xlsx",
                Filter = "Excel файлы (Spisok.xlsx)|*.xlsx",
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
            for (int i = 0; i < _rows; i++)
            {
                for (int j = 0; j < _columns; j++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjWorkExcel);
            ObjWorkSheet = null;
            ObjWorkBook = null;
            ObjWorkExcel = null;
            GC.Collect();

            using (ИСРПОEntities2 иСРПОEntities = new ИСРПОEntities2())
            {
                иСРПОEntities.User.RemoveRange(иСРПОEntities.User);
                иСРПОEntities.SaveChanges();
            }


            using (ИСРПОEntities2 иСРПОEntities = new ИСРПОEntities2())
            {
                for (int i = 1; i < _rows; i++)
                {
                    иСРПОEntities.User.Add(new User() { ID = Convert.ToInt32(list[i,0]), Наименование_услуги = list[i, 1], Вид_услуги = list[i, 2], Стоимость = list[i, 3] });
                }
                иСРПОEntities.SaveChanges();
            }

        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<User> all;
            using (ИСРПОEntities2 иСРПОEntities = new ИСРПОEntities2())
            {
                all = иСРПОEntities.User.ToList();
            }

            var categorizedItems = all
                .Select(u =>
                    new
                    {
                        User = u,
                        Cost = double.Parse(u.Стоимость)
                    })
                .OrderBy(u => u.Cost)
                .Select(u =>
                {
                    if (u.Cost <= 350)
                        return new { Category = "Категория 1", User = u.User };
                    else if (u.Cost > 250 && u.Cost <= 800)
                        return new { Category = "Категория 2", User = u.User };
                    else
                        return new { Category = "Категория 3", User = u.User };
                })
                .GroupBy(x => x.Category)
                .ToList();

            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add();
            int startRowIndex = 1;

            foreach (var category in categorizedItems)
            {
                Excel.Worksheet worksheet = app.Worksheets.Add();
                worksheet.Name = category.Key;
                worksheet.Cells[1][startRowIndex] = "Наименование услуги";
                worksheet.Cells[2][startRowIndex] = "Стоимость";
                startRowIndex++;

                foreach (var item in category)
                {
                    worksheet.Cells[1][startRowIndex] = item.User.Наименование_услуги;
                    worksheet.Cells[2][startRowIndex] = item.User.Стоимость;
                    startRowIndex++;
                }

                Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, startRowIndex - 1]];
                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Columns.AutoFit();
                startRowIndex = 1;
            }

            app.Visible = true;



        }

        
    }
}
