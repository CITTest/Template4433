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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4333
{
    /// <summary>
    /// Логика взаимодействия для Emel_4333_3lab.xaml
    /// </summary>
    public partial class Emel_4333_3lab : System.Windows.Window
    {
        public Emel_4333_3lab()
        {
            InitializeComponent();
        }
        private void BnImport_Click(object sender,
RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (1.xlsx)|*.xlsx",
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
            using (serviceEntities usersEntities = new serviceEntities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.s_ervice.Add(new s_ervice()
                    {
                        name_service = list[i, 1],
                        type_of_service = list[i, 2],
                        code_service = list[i, 3],
                        price = Convert.ToInt32(list[i, 4])
                    });
                }
                usersEntities.SaveChanges();
            }
        }
        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<s_ervice> alldannie;
            List<string> strings;
            using (serviceEntities SEREntities = new serviceEntities())
            {
                alldannie =
                SEREntities.s_ervice.OrderBy(s => s.price).ToList();
                strings =
                SEREntities.s_ervice.ToList().Select(s_ervice => s_ervice.type_of_service.ToString()).Distinct().ToList();
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
                worksheet.Cells[1][startRowIndex] = "id";
                worksheet.Cells[2][startRowIndex] = "Название услуги";
                worksheet.Cells[3][startRowIndex] = "стоимость";
                startRowIndex++;
                foreach (var dann in alldannie)
                {

                        if (dann.type_of_service == strings[i])
                        {
                            worksheet.Cells[1][startRowIndex] = dann.id;
                            worksheet.Cells[2][startRowIndex] = dann.name_service;
                            worksheet.Cells[3][startRowIndex] = dann.price;
                            startRowIndex++;
                        }
                }
            }
            app.Visible = true;
        }
    }

}
