// ЛР №25, пункты 3.1 и 3.2. WPF (.NET Framework), решение LabWork22.
// Ссылки: System.Data, COM -> Microsoft Excel 16.0 Object Library.
// Замените MainWindow.xaml и MainWindow.xaml.cs этими файлами.
// В App.xaml должно быть StartupUri="MainWindow.xaml".
// БД ispp5503: кинотеатр - Hall.Cinema, номер зала - Hall.HallId.

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace LabWork22
{
    public partial class MainWindow : Window
    {
        // Поменяйте Data Source на свой сервер
        const string Conn = @"Data Source=.\SQLEXPRESS;Initial Catalog=ispp5503;Integrated Security=True";

        public MainWindow()
        {
            InitializeComponent();
            datePicker.SelectedDate = DateTime.Today;
        }

        void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Run(() => cinemaBox.ItemsSource =
                Query("SELECT DISTINCT Cinema FROM Hall ORDER BY Cinema", null)
                    .Rows.Cast<DataRow>().Select(r => (string)r[0]).ToList());
        }

        void Run(Action action)
        {
            try { action(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Ошибка"); }
        }

        // Все данные берутся из БД
        DataTable Query(string sql, DateTime? date)
        {
            var table = new DataTable();
            using (var adapter = new SqlDataAdapter(sql, Conn))
            {
                if (date != null)
                    adapter.SelectCommand.Parameters.AddWithValue("@d", date.Value.Date);
                adapter.Fill(table);
            }
            return table;
        }

        // Расписание на день, сортировка по фильму и времени начала
        DataTable Schedule()
        {
            DateTime date = datePicker.SelectedDate ?? DateTime.Today;
            return Query(@"SELECT h.Cinema, f.Title, s.StartDateTime, h.HallId, s.Price
                           FROM Session s
                           JOIN Film f ON f.Id = s.Film_Id
                           JOIN Hall h ON h.Id = s.Hall_Id
                           WHERE CAST(s.StartDateTime AS date) = @d
                           ORDER BY f.Title, s.StartDateTime", date);
        }

        void CsvButton_Click(object sender, RoutedEventArgs e) { Run(SaveCsv); }
        void XlsxButton_Click(object sender, RoutedEventArgs e) { Run(SaveXlsx); }

        // 3.1 csv: расписание выбранного кинотеатра
        void SaveCsv()
        {
            string cinema = cinemaBox.SelectedItem as string;
            if (cinema == null) { MessageBox.Show("Выберите кинотеатр."); return; }

            var rows = Schedule().Rows.Cast<DataRow>()
                .Where(r => (string)r["Cinema"] == cinema).ToList();
            if (rows.Count == 0) { MessageBox.Show("Сеансов нет."); return; }

            var dialog = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "Расписание.csv" };
            if (dialog.ShowDialog() != true) return;

            // StreamWriter, UTF-8 с BOM, разделитель ';' (чтобы Excel открыл русский текст)
            using (var writer = new StreamWriter(dialog.FileName, false, new UTF8Encoding(true)))
            {
                writer.WriteLine("Фильм;Начало сеанса;Зал;Цена");
                foreach (var r in rows)
                    writer.WriteLine($"{r["Title"]};{((DateTime)r["StartDateTime"]):HH:mm};{r["HallId"]};{r["Price"]}");
            }
            MessageBox.Show("Файл сохранён.");
        }

        // 3.2 xlsx: каждый кинотеатр на отдельном листе (название листа = название кинотеатра)
        void SaveXlsx()
        {
            var all = Schedule().Rows.Cast<DataRow>().ToList();
            if (all.Count == 0) { MessageBox.Show("Сеансов нет."); return; }

            var dialog = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "Расписание.xlsx" };
            if (dialog.ShowDialog() != true) return;

            var cinemas = all.Select(r => (string)r["Cinema"]).Distinct().OrderBy(c => c).ToList();

            var excelApp = new Excel.Application { DisplayAlerts = false };
            try
            {
                excelApp.SheetsInNewWorkbook = cinemas.Count;   // по листу на кинотеатр
                var workbook = excelApp.Workbooks.Add();

                for (int i = 0; i < cinemas.Count; i++)
                {
                    Excel.Worksheet sheet = workbook.Worksheets[i + 1];
                    // в названии листа нельзя \ / ? * [ ] : и больше 31 символа
                    string name = new string(cinemas[i].Where(c => "\\/?*[]:".IndexOf(c) < 0).ToArray());
                    sheet.Name = name.Length > 31 ? name.Substring(0, 31) : name;

                    string[] head = { "Фильм", "Начало сеанса", "Зал", "Цена" };
                    for (int c = 0; c < head.Length; c++) sheet.Cells[1, c + 1] = head[c];
                    sheet.Range["A1:D1"].Font.Bold = true;

                    int row = 2;
                    foreach (var r in all.Where(x => (string)x["Cinema"] == cinemas[i]))
                    {
                        sheet.Cells[row, 1] = r["Title"].ToString();
                        sheet.Cells[row, 2] = ((DateTime)r["StartDateTime"]).ToString("HH:mm");
                        sheet.Cells[row, 3] = Convert.ToInt32(r["HallId"]);
                        sheet.Cells[row, 4] = Convert.ToDecimal(r["Price"]);
                        row++;
                    }
                    sheet.Range["A1:D" + (row - 1)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    sheet.Columns.AutoFit();
                }

                if (File.Exists(dialog.FileName)) File.Delete(dialog.FileName);
                workbook.SaveAs(dialog.FileName);
                MessageBox.Show("Файл сохранён.");
            }
            finally { excelApp.Quit(); }
        }
    }
}
