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
            string GetCategory(int cost)
            {
                if (cost < 350)
                    return "0-350";
                else if (cost >= 350 && cost < 800)
                    return "350-800";
                else
                    return "800+";
            }

        }
        }
    }

