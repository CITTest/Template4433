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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
namespace Template4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Borodin.xaml
    /// </summary>
    public partial class _4333_Borodin : Window
    {
        public _4333_Borodin()
        {
            InitializeComponent();
        }
        private void ImportClick(object sender, RoutedEventArgs e)
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
            using (isrpoEntities2 isrpoEntities = new isrpoEntities2())
            {
                for (int i = 0; i < _rows; i++)
                {
                    if (list[i, 1] != "" && list[i, 2] != "" && list[i, 3] != "" && list[i, 4] != "")
                    {


                        isrpoEntities.Table_1.Add(new Table_1()
                        {
                            OrderCode = list[i, 1],
                            DateCreation = list[i, 2],
                            OrderTime = list[i, 3],
                            ClientCode = list[i, 4],
                            Services = list[i, 5],
                            Status = list[i, 6],
                            ClosingDate = list[i, 7],
                            RentalTime = list[i, 8]
                        });
                    }

                }
                isrpoEntities.SaveChanges();
            }
        }

        private void ExportClick(object sender, RoutedEventArgs e)
        {
            /*List<Table_1> allRentalTime;
            List<Table_1> allOrderCode;
            using (isrpoEntities2 isrpoEntities = new isrpoEntities2())
            {

                allRentalTime = isrpoEntities.Table_1.ToList().GroupBy(x => x.RentalTime).Select(y => y.First()).ToList();
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = allRentalTime.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                for (int i = 0; i < allRentalTime.Count(); i++)
                {
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = Convert.ToString(allRentalTime[i].OrderCode);
                    worksheet.Cells[1][2] = "Время проката";
                    worksheet.Cells[2][2] = "Код Заказа";
                    startRowIndex++;
                    var OrderCodeCategories = allOrderCode.GroupBy(s => s.OrderCode).ToList();
                    foreach (var students in studentsCategories)
                    {
                        if (students.Key == allGroups[i].Id)
                        {
                            Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                            headerRange.Merge();
                            headerRange.Value = allGroups[i].NumberGroup;
                            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            headerRange.Font.Italic = true;
                            startRowIndex++;
                            foreach (Student student in allStudents)
                            {
                                if (student.NumberGroupId == students.Key)
                                {
                                    worksheet.Cells[1][startRowIndex] = student.Id;
                                    worksheet.Cells[2][startRowIndex] = student.Name;
                                    startRowIndex++;
                                }
                            }
                            worksheet.Cells[1][startRowIndex].Formula = $"=СЧЁТ(A3:A{startRowIndex - 1})";
                            worksheet.Cells[1][startRowIndex].Font.Bold = true;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }*/
            List<Table_1> allorders;
            List<string> status;
            using (isrpoEntities2 isrpoEntities = new isrpoEntities2())
            {
                allorders = isrpoEntities.Table_1.ToList().OrderBy(s => s.Status).ToList();
                status = isrpoEntities.Table_1.ToList().Select(Ord => Ord.Status.ToString()).Distinct().ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = status.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < status.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(status[i]);
                worksheet.Cells[1][startRowIndex] = "ID";
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][startRowIndex] = "Дата создания";
                worksheet.Cells[4][startRowIndex] = "Код клиента";
                worksheet.Cells[5][startRowIndex] = "Услуги";
                startRowIndex++;
                foreach (var order in allorders)
                {
                    if (order.Status == status[i])
                    {
                        worksheet.Cells[1][startRowIndex] = order.Id.ToString();
                        worksheet.Cells[2][startRowIndex] = order.OrderCode;
                        worksheet.Cells[3][startRowIndex] = order.DateCreation;
                        worksheet.Cells[4][startRowIndex] = order.ClientCode;
                        worksheet.Cells[5][startRowIndex] = order.Services;
                        startRowIndex++;
                    }
                }

            }
            app.Visible = true;
        }
    }
}
