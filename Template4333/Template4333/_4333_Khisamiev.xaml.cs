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
	/// Логика взаимодействия для _4333_Khisamiev.xaml
	/// </summary>
	public partial class _4333_Khisamiev : Window
	{
		public _4333_Khisamiev()
		{
			InitializeComponent();
		}

		private void Import_Click(object sender, RoutedEventArgs e)
		{

			OpenFileDialog ofd = new OpenFileDialog()
			{
				DefaultExt = "*.xls;*.xlsx",
				Filter = "файл Excel (3.xlsx)|*.xlsx",
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
			using (ISRPO3Entities isproEntities = new ISRPO3Entities())
			{
				for (int i = 1; i < _rows; i++)
				{
					isproEntities.ISRPO3.Add(new ISRPO3()
					{
						FIO = list[i, 0],
						ClienID = list[i, 1],
						DateBirth = list[i, 2],
						Indeks = list[i, 3],
						City = list[i, 4],
						Ulitca = list[i, 5],
						Home = list[i, 6],
						Kvartira = list[i, 7],
						Email = list[i, 8],

					}); ;
				}
				isproEntities.SaveChanges();
			}
			BnImport.Background = new SolidColorBrush(Colors.Green);
			BnImport.Content = "Импорт выполнен успешно!";
		}

		private void Export_Click(object sender, RoutedEventArgs e)
		{
			List<ISRPO3> all;
			List<string> afterSort;
			using (ISRPO3Entities iSRPO3Entities = new ISRPO3Entities())
			{
				all = iSRPO3Entities.ISRPO3.OrderBy(s => s.Ulitca).ToList();
				afterSort = iSRPO3Entities.ISRPO3.ToList().Select(ISRPO3 => ISRPO3.Ulitca.ToString()).Distinct().ToList();
			}
			var app = new Excel.Application();
			app.SheetsInNewWorkbook = afterSort.Count();
			Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
			for (int i = 1; i < afterSort.Count(); i++)
			{
				int startRowIndex = 1;
				Excel.Worksheet worksheet = app.Worksheets.Item[i];
				worksheet.Name = afterSort[i];
				worksheet.Cells[1][startRowIndex] = "Код клиента";
				worksheet.Cells[2][startRowIndex] = "ФИО";
				worksheet.Cells[3][startRowIndex] = "Email";
				startRowIndex++;
				foreach (var newTable in all)
				{
					if (newTable.Ulitca == afterSort[i])
					{
						worksheet.Cells[1][startRowIndex] = newTable.ClienID;
						worksheet.Cells[2][startRowIndex] = newTable.FIO;
						worksheet.Cells[3][startRowIndex] = newTable.Email;
						startRowIndex++;
						Excel.Range range = worksheet.Range["A2:C10"];
						range.Sort(range.Columns[2]);
					}

				}

			}
			app.Visible = true;
			BnExport.Background = new SolidColorBrush(Colors.Green);
			BnExport.Content = "Экспорт выполнен успешно!";
			GC.Collect();
		}
	}
}
