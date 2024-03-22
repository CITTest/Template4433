using LanguageExt.TypeClasses;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Serialization;
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
using static Template4333.MainWindow;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;




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

        //public class User
        //{
        //    [JsonPropertyName("IdServices")]
        //    public int ID { get; set; }
        //    [JsonPropertyName("NameServices")]
        //    public string Наименование_услуги { get; set; }
        //    [JsonPropertyName("TypeOfService")]
        //    public string Вид_услуги { get; set; }
        //    [JsonPropertyName("Cost")]
        //    public double Стоимость { get; set; }
        //}



        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<User> all;
            using (ИСРПОEntities4 иСРПОEntities = new ИСРПОEntities4())
            {
                all = иСРПОEntities.User.ToList();
            }

            var categorizedItems = all
                .Select(u =>
                    new
                    {
                        User = u,
                        Cost = u.Стоимость
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
            Word.Application app = new Word.Application();
            app.Visible = true;


            Word.Document doc = app.Documents.Add();


            bool isFirstGroup = true;

            foreach (var PlusGroup in categorizedItems)
            {

                if (!isFirstGroup)
                {
                    Word.Range rng = doc.Content;
                    rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    rng.InsertBreak(Word.WdBreakType.wdPageBreak);
                }


                doc.Content.Text += PlusGroup.Key + "\n";
                doc.Content.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;


                foreach (var item in PlusGroup)
                {
                    string position = item.User.Наименование_услуги;
                    int login = item.User.Стоимость;

                    doc.Content.Text += $"Наименование услуги: {position}, Стоимость: {login}\n";
                }

                isFirstGroup = false;


            }
        }


        private async void BnImport_Click(object sender, RoutedEventArgs e)
        {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "JSON files (*.json)|*.json";

                if (openFileDialog.ShowDialog() == true)
                {
                    string jsonFilePath = openFileDialog.FileName;

                    List<User> ordersD;

                    using (FileStream fs = new FileStream(jsonFilePath, FileMode.Open))
                    {
                        ordersD = await System.Text.Json.JsonSerializer.DeserializeAsync<List<User>>(fs);
                    }

                using (ИСРПОEntities4 иСРПОEntities = new ИСРПОEntities4())
                {
                    иСРПОEntities.User.RemoveRange(иСРПОEntities.User);
                    иСРПОEntities.SaveChanges();
                }

                using (ИСРПОEntities4 usersEntities = new ИСРПОEntities4())
                    {
                        foreach (var orderD in ordersD)
                        {
                            usersEntities.User.Add(new User()
                            {
                                ID = orderD.ID,
                                Наименование_услуги = orderD.Наименование_услуги,
                                Вид_услуги = orderD.Вид_услуги,
                                Стоимость = orderD.Стоимость
                            });
                        }
                        usersEntities.SaveChanges();
                    }

                    MessageBox.Show("Данные успешно импортированы из JSON файла в таблицу БД.");
                }
            


        }
    }

   
}
