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
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;
using System.Security.Cryptography;

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

		private void ExportWord_Click(object sender, RoutedEventArgs e)
		{
			List<ISRPO3JSON> all;

			using (ISRPO3Entities1 usersEntities = new ISRPO3Entities1())
			{
				all = usersEntities.ISRPO3JSON.ToList().OrderBy(s => s.Street).ToList();

			}
			foreach (var group in all.GroupBy(o => o.Street))
			{
				var app = new Word.Application();
				Word.Document document = app.Documents.Add();

				Word.Paragraph headerParagraph = document.Paragraphs.Add();
				Word.Range headerRange = headerParagraph.Range;
				headerRange.Text = $"Улица: {group.Key}";
				headerParagraph.set_Style("Заголовок 1");
				headerRange.InsertParagraphAfter();

				Word.Table mytable = document.Tables.Add(headerRange, group.Count() + 1, 3);
				mytable.Borders.InsideLineStyle = mytable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

				mytable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

				mytable.Cell(1, 1).Range.Text = "Код клиента";
				mytable.Cell(1, 2).Range.Text = "ФИО";
				mytable.Cell(1, 3).Range.Text = "E-mail";

				int i = 1;
				foreach (var by in group.OrderBy(s => s.FullName))
				{
					i++;
					mytable.Cell(i, 1).Range.Text = by.CodeClient;
					mytable.Cell(i, 2).Range.Text = by.FullName;
					mytable.Cell(i, 3).Range.Text = by.E_mail;
				}

				string fileName = $"C:/Users/Puchindoo/Desktop/ISRPO4/{group.Key}.docx";
				BnExportWord.Background = new SolidColorBrush(Colors.Green);
				BnExport.Content = "Экспорт выполнен успешно!";
				app.Visible = true;
			}
		}

		private async void ImortJSON_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "JSON files (*.json)|*.json";

			if (openFileDialog.ShowDialog() == true)
			{
				string jsonFilePath = openFileDialog.FileName;

				List<ISRPO3JSON> all;

				using (FileStream fs = new FileStream(jsonFilePath, FileMode.Open))
				{
					all = await JsonSerializer.DeserializeAsync<List<ISRPO3JSON>>(fs);
				}

				using (ISRPO3Entities1 isrpo3Entities = new ISRPO3Entities1())
				{
					foreach (var alldata in all)
					{
						ISRPO3JSON mytable = new ISRPO3JSON
						{
							FullName = alldata.FullName,
							CodeClient = alldata.CodeClient,
							BirthDate = alldata.BirthDate,
							Index = alldata.Index,
							City = alldata.City,
							Street = alldata.Street,
							Home = alldata.Home,
							Kvartira = alldata.Kvartira,
							E_mail = alldata.E_mail,

						};
						isrpo3Entities.ISRPO3JSON.Add(mytable);
					}
					isrpo3Entities.SaveChanges();
					BnImportJSON.Content = "Импорт выполнен успешно!";
					BnImportJSON.Background = new SolidColorBrush(Colors.Green);
				}

			}
		}
	}
}
