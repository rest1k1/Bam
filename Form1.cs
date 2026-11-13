// ЛР №25, пункты 3.1 и 3.2. Windows Forms (.NET Framework), решение LabWork22.
// Ссылки: System.Data, System.Windows.Forms, System.Drawing,
//         COM -> Microsoft Excel 16.0 Object Library.
// Форма создаётся кодом: вставьте этот файл вместо Form1.cs и удалите Form1.Designer.cs.
// БД ispp5503: кинотеатр - Hall.Cinema, номер зала - Hall.HallId.

using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LabWork22
{
    public class Form1 : Form
    {
        // Поменяйте Data Source на свой сервер
        const string Conn = @"Data Source=.\SQLEXPRESS;Initial Catalog=ispp5503;Integrated Security=True";

        DateTimePicker datePicker = new DateTimePicker { Left = 115, Top = 15, Width = 300, Format = DateTimePickerFormat.Short };
        ComboBox cinemaBox = new ComboBox { Left = 115, Top = 55, Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };

        public Form1()
        {
            Text = "Экспорт расписания";
            ClientSize = new System.Drawing.Size(440, 190);
            StartPosition = FormStartPosition.CenterScreen;
            var csvBtn = new Button { Text = "3.1 Сохранить в csv (выбранный кинотеатр)", Left = 15, Top = 100, Width = 400, Height = 30 };
            var xlsxBtn = new Button { Text = "3.2 Сохранить в xlsx (все кинотеатры)", Left = 15, Top = 140, Width = 400, Height = 30 };
            csvBtn.Click += (s, e) => Run(SaveCsv);
            xlsxBtn.Click += (s, e) => Run(SaveXlsx);
            Controls.AddRange(new Control[] {
                new Label { Text = "Дата:", Left = 15, Top = 19 },
                new Label { Text = "Кинотеатр:", Left = 15, Top = 59 },
                datePicker, cinemaBox, csvBtn, xlsxBtn });
            Load += (s, e) => Run(() =>
                cinemaBox.DataSource = Query("SELECT DISTINCT Cinema FROM Hall ORDER BY Cinema", null)
                    .Rows.Cast<DataRow>().Select(r => (string)r[0]).ToList());
        }

        // Выполнение действия с показом ошибок
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

        // Расписание на день: кинотеатр, фильм, начало, зал, цена.
        // Сортировка по фильму и времени начала.
        DataTable Schedule(DateTime date)
        {
            return Query(@"SELECT h.Cinema, f.Title, s.StartDateTime, h.HallId, s.Price
                           FROM Session s
                           JOIN Film f ON f.Id = s.Film_Id
                           JOIN Hall h ON h.Id = s.Hall_Id
                           WHERE CAST(s.StartDateTime AS date) = @d
                           ORDER BY f.Title, s.StartDateTime", date);
        }

        // 3.1 csv: расписание выбранного кинотеатра
        void SaveCsv()
        {
            string cinema = cinemaBox.Text;
            var rows = Schedule(datePicker.Value).Rows.Cast<DataRow>()
                .Where(r => (string)r["Cinema"] == cinema).ToList();
            if (rows.Count == 0) { MessageBox.Show("Сеансов нет."); return; }

            var dialog = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "Расписание.csv" };
            if (dialog.ShowDialog() != DialogResult.OK) return;

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
            var all = Schedule(datePicker.Value).Rows.Cast<DataRow>().ToList();
            if (all.Count == 0) { MessageBox.Show("Сеансов нет."); return; }

            var dialog = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "Расписание.xlsx" };
            if (dialog.ShowDialog() != DialogResult.OK) return;

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
