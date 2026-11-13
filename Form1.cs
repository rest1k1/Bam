// ЛР №25. Экспорт расписания в csv / xlsx / html.
// Проект: Windows Forms (.NET Framework), решение LabWork22.
// Ссылки (References): System.Data, System.Windows.Forms, System.Drawing,
//   Microsoft.Office.Interop.Excel  (COM: Microsoft Excel Object Library).
// Форма создаётся кодом (без дизайнера): замените содержимое Form1.cs этим файлом
// (в Form1.Designer.cs оставьте только пустой InitializeComponent() или удалите файл).
// Шаблон: положите "Шаблон для фильма.xlsx" в папку Шаблоны рядом с exe
// (Свойства файла -> Копировать в выходной каталог: всегда) и назовите его
// "Шаблон_для_фильма.xlsx" либо поправьте константу TemplatePath ниже.
//
// ВАЖНО: схему БД и строку подключения подправьте под свою БД (раздел "Работа с БД").
// Используются таблицы БД ispp5503: Film(Id, Title, Duration), Hall(Id, Cinema, HallId),
//   Session(Id, Film_Id, Hall_Id, Price, StartDateTime). Название кинотеатра - Hall.Cinema,
//   номер зала - Hall.HallId. Строку подключения (Data Source) поправьте под свой сервер.

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LabWork22
{
    // Строка расписания
    public class ScheduleRow
    {
        public int CinemaId { get; set; }
        public string Cinema { get; set; }
        public int FilmId { get; set; }
        public string Film { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public int Hall { get; set; }
        public decimal Price { get; set; }
    }

    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public override string ToString() { return Name; }
    }

    public class Form1 : Form
    {
        // ---------- настройки ----------
        private const string ConnectionString =
            @"Data Source=.\SQLEXPRESS;Initial Catalog=ispp5503;Integrated Security=True";
        private static readonly string TemplatePath =
            Environment.CurrentDirectory + "\\Шаблоны\\Шаблон_для_фильма.xlsx";
        private static readonly CultureInfo Ru = new CultureInfo("ru-RU");

        // ---------- элементы управления ----------
        private ComboBox cinemaBox = new ComboBox();
        private ComboBox filmBox = new ComboBox();
        private DateTimePicker datePicker = new DateTimePicker();
        private Button csvButton = new Button();
        private Button xlsxButton = new Button();
        private Button templateButton = new Button();
        private Button saveButton = new Button();

        public Form1()
        {
            Text = "Экспорт расписания";
            ClientSize = new Size(440, 300);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            AddRow("Дата:", datePicker, 15);
            AddRow("Кинотеатр:", cinemaBox, 55);
            AddRow("Фильм:", filmBox, 95);
            cinemaBox.DropDownStyle = ComboBoxStyle.DropDownList;
            filmBox.DropDownStyle = ComboBoxStyle.DropDownList;
            datePicker.Format = DateTimePickerFormat.Short;

            AddButton(csvButton, "3.1 Сохранить в csv (выбранный кинотеатр)", 140, csvButton_Click);
            AddButton(xlsxButton, "3.2 Сохранить в xlsx (все кинотеатры, листы)", 175, xlsxButton_Click);
            AddButton(templateButton, "3.3 Сохранить в xlsx по шаблону (выбранный фильм)", 210, templateButton_Click);
            AddButton(saveButton, "3.4 Сохранить расписание (xlsx / html)", 245, saveButton_Click);

            Load += Form1_Load;
        }

        private void AddRow(string caption, Control control, int top)
        {
            Controls.Add(new Label { Text = caption, Left = 15, Top = top + 4, Width = 90 });
            control.Left = 115; control.Top = top; control.Width = 300;
            Controls.Add(control);
        }

        private void AddButton(Button b, string text, int top, EventHandler handler)
        {
            b.Text = text; b.Left = 15; b.Top = top; b.Width = 400; b.Height = 30;
            b.Click += handler;
            Controls.Add(b);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                cinemaBox.DataSource = LoadCinemas();
                filmBox.DataSource = LoadItems("SELECT Id, Title FROM Film ORDER BY Title");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось загрузить данные из БД:\n" + ex.Message,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =====================================================================
        //                           Работа с БД
        // =====================================================================
        private List<Item> LoadItems(string sql)
        {
            var list = new List<Item>();
            using (var con = new SqlConnection(ConnectionString))
            using (var cmd = new SqlCommand(sql, con))
            {
                con.Open();
                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        list.Add(new Item { Id = r.GetInt32(0), Name = r.GetString(1) });
            }
            return list;
        }

        // Кинотеатры - уникальные значения Hall.Cinema (отдельной таблицы нет)
        private List<Item> LoadCinemas()
        {
            var list = new List<Item>();
            using (var con = new SqlConnection(ConnectionString))
            using (var cmd = new SqlCommand("SELECT DISTINCT Cinema FROM Hall ORDER BY Cinema", con))
            {
                con.Open();
                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        list.Add(new Item { Id = list.Count + 1, Name = r.GetString(0) });
            }
            return list;
        }

        // Сеансы на день; cinemaId / filmId = null -> без фильтра.
        // Сортировка: по фильму, затем по времени начала.
        private List<ScheduleRow> LoadSchedule(DateTime date, string cinemaName, int? filmId)
        {
            const string sql = @"
SELECT h.Cinema, f.Id, f.Title, s.StartDateTime, f.Duration, h.HallId, s.Price
FROM Session s
JOIN Film f ON f.Id = s.Film_Id
JOIN Hall h ON h.Id = s.Hall_Id
WHERE s.StartDateTime >= @d1 AND s.StartDateTime < @d2
  AND (@cinema IS NULL OR h.Cinema = @cinema)
  AND (@film IS NULL OR f.Id = @film)
ORDER BY f.Title, s.StartDateTime";

            var list = new List<ScheduleRow>();
            using (var con = new SqlConnection(ConnectionString))
            using (var cmd = new SqlCommand(sql, con))
            {
                cmd.Parameters.AddWithValue("@d1", date.Date);
                cmd.Parameters.AddWithValue("@d2", date.Date.AddDays(1));
                cmd.Parameters.AddWithValue("@cinema", (object)cinemaName ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@film", (object)filmId ?? DBNull.Value);
                con.Open();
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        var start = r.GetDateTime(3);
                        list.Add(new ScheduleRow
                        {
                            CinemaId = 0,
                            Cinema = r.GetString(0),
                            FilmId = r.GetInt32(1),
                            Film = r.GetString(2),
                            Start = start,
                            End = start.AddMinutes(Convert.ToDouble(r.GetValue(4))),
                            Hall = Convert.ToInt32(r.GetValue(5)),
                            Price = Convert.ToDecimal(r.GetValue(6))
                        });
                    }
                }
            }
            return list;
        }

        // =====================================================================
        //                       Вспомогательные методы
        // =====================================================================
        private static string AskSavePath(string filter, string defaultName)
        {
            using (var dlg = new SaveFileDialog { Filter = filter, FileName = defaultName })
                return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
        }

        private static string CsvEscape(string s)
        {
            if (s.Contains(";") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        // Имя листа Excel: до 31 символа, без \ / ? * [ ] :
        private static string SheetName(string name, HashSet<string> used)
        {
            var clean = new string(name.Where(ch => "\\/?*[]:".IndexOf(ch) < 0).ToArray()).Trim();
            if (clean.Length == 0) clean = "Лист";
            if (clean.Length > 31) clean = clean.Substring(0, 31);
            string result = clean; int n = 2;
            while (!used.Add(result.ToLower()))
            {
                string suffix = " (" + n++ + ")";
                result = clean.Substring(0, Math.Min(clean.Length, 31 - suffix.Length)) + suffix;
            }
            return result;
        }

        private static void ReleaseCom(object o)
        {
            try { if (o != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(o); } catch { }
        }

        private static void Header(Excel.Worksheet ws, int row, params string[] titles)
        {
            for (int i = 0; i < titles.Length; i++) ws.Cells[row, i + 1] = titles[i];
            Excel.Range r = ws.Range[ws.Cells[row, 1], ws.Cells[row, titles.Length]];
            r.Font.Bold = true;
            r.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        // =====================================================================
        //  3.1  csv: расписание выбранного кинотеатра на выбранный день
        // =====================================================================
        private void csvButton_Click(object sender, EventArgs e)
        {
            try
            {
                var cinema = cinemaBox.SelectedItem as Item;
                if (cinema == null) { MessageBox.Show("Выберите кинотеатр."); return; }

                var rows = LoadSchedule(datePicker.Value, cinema.Name, null)
                    .OrderBy(r => r.Film).ThenBy(r => r.Start).ToList();
                if (rows.Count == 0) { MessageBox.Show("Сеансов на выбранный день нет."); return; }

                string path = AskSavePath("CSV (*.csv)|*.csv", "Расписание.csv");
                if (path == null) return;

                var sb = new StringBuilder();
                sb.AppendLine("Фильм;Начало сеанса;Зал;Цена");
                foreach (var r in rows)
                    sb.AppendLine(string.Join(";", CsvEscape(r.Film), r.Start.ToString("HH:mm"),
                        r.Hall, r.Price.ToString("0.00", Ru)));

                // UTF-8 с BOM и ';' - чтобы русский текст корректно открывался в Excel
                File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
                MessageBox.Show("Файл сохранён.");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        // =====================================================================
        //  3.2  xlsx: каждый кинотеатр на отдельном листе
        // =====================================================================
        private void xlsxButton_Click(object sender, EventArgs e)
        {
            try
            {
                string path = AskSavePath("Excel (*.xlsx)|*.xlsx", "Расписание.xlsx");
                if (path == null) return;
                SaveAllCinemasXlsx(path, datePicker.Value);
                MessageBox.Show("Файл сохранён.");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void SaveAllCinemasXlsx(string path, DateTime date)
        {
            var all = LoadSchedule(date, null, null);
            if (all.Count == 0) throw new InvalidOperationException("Сеансов на выбранный день нет.");

            var groups = all.GroupBy(r => new { r.CinemaId, r.Cinema })
                            .OrderBy(g => g.Key.Cinema).ToList();

            var app = new Excel.Application { DisplayAlerts = false };
            Excel.Workbook wb = null;
            try
            {
                app.SheetsInNewWorkbook = groups.Count;     // по листу на кинотеатр
                wb = app.Workbooks.Add();
                var used = new HashSet<string>();

                for (int i = 0; i < groups.Count; i++)
                {
                    Excel.Worksheet ws = wb.Worksheets[i + 1];
                    ws.Name = SheetName(groups[i].Key.Cinema, used);
                    Header(ws, 1, "Фильм", "Начало сеанса", "Зал", "Цена");

                    int row = 2;
                    foreach (var r in groups[i].OrderBy(x => x.Film).ThenBy(x => x.Start))
                    {
                        ws.Cells[row, 1] = r.Film;
                        ws.Cells[row, 2] = r.Start.ToString("HH:mm");
                        ws.Cells[row, 3] = r.Hall;
                        ws.Cells[row, 4] = r.Price;
                        row++;
                    }
                    Excel.Range table = ws.Range[ws.Cells[1, 1], ws.Cells[row - 1, 4]];
                    table.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    ws.Range[ws.Cells[2, 2], ws.Cells[row - 1, 3]].HorizontalAlignment =
                        Excel.XlHAlign.xlHAlignCenter;
                    ws.Range[ws.Cells[2, 4], ws.Cells[row - 1, 4]].NumberFormat = "0.00 \"руб\"";
                    ws.Columns.AutoFit();
                }

                if (File.Exists(path)) File.Delete(path);
                wb.SaveAs(path);
            }
            finally
            {
                if (wb != null) { wb.Close(false); ReleaseCom(wb); }
                app.Quit(); ReleaseCom(app);
            }
        }

        // =====================================================================
        //  3.3  xlsx по шаблону: выбранный фильм на выбранный день
        //  Шаблон: A1 - название фильма (объединено A1:W1), A3 - дата,
        //  далее блоки по 5 строк на кинотеатр, начиная со строки 5:
        //    строка N   - название кинотеатра
        //    строка N+1 - "hh:mm - hh:mm"   (сеансы идут по столбцам A, B, C ...)
        //    строка N+2 - "Зал N"
        //    строка N+3 - "N руб"
        //    строка N+4 - пустая
        // =====================================================================
        private void templateButton_Click(object sender, EventArgs e)
        {
            try
            {
                var film = filmBox.SelectedItem as Item;
                if (film == null) { MessageBox.Show("Выберите фильм."); return; }
                string path = AskSavePath("Excel (*.xlsx)|*.xlsx", "Расписание фильма.xlsx");
                if (path == null) return;
                SaveFilmByTemplate(path, datePicker.Value, film);
                MessageBox.Show("Файл сохранён.");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void SaveFilmByTemplate(string path, DateTime date, Item film)
        {
            if (!File.Exists(TemplatePath))
                throw new FileNotFoundException("Не найден файл шаблона:\n" + TemplatePath);

            var rows = LoadSchedule(date, null, film.Id);
            if (rows.Count == 0) throw new InvalidOperationException("Сеансов выбранного фильма на этот день нет.");

            // группировка по кинотеатру, сортировка по времени начала
            var groups = rows.GroupBy(r => r.Cinema).OrderBy(g => g.Key).ToList();

            const int firstBlockRow = 5, blockHeight = 5, lastCol = 23; // A..W
            var app = new Excel.Application { DisplayAlerts = false };
            Excel.Workbook wb = null;
            try
            {
                wb = app.Workbooks.Open(TemplatePath);
                Excel.Worksheet ws = wb.Worksheets[1];

                // чистим блоки-заготовки шаблона (кроме форматирования)
                ws.Range[ws.Cells[firstBlockRow, 1], ws.Cells[firstBlockRow + 3 * blockHeight - 1, lastCol]]
                    .ClearContents();

                ws.Cells[1, 1] = film.Name;
                ws.Cells[3, 1] = date.ToString("d MMMM", Ru);

                for (int i = 0; i < groups.Count; i++)
                {
                    int top = firstBlockRow + i * blockHeight;

                    // если кинотеатров больше, чем блоков в шаблоне (3) - копируем оформление первого блока
                    if (i >= 3)
                    {
                        Excel.Range src = ws.Range[ws.Cells[firstBlockRow, 1],
                                                   ws.Cells[firstBlockRow + blockHeight - 1, lastCol]];
                        src.Copy(ws.Cells[top, 1]);
                        ws.Range[ws.Cells[top, 1], ws.Cells[top + blockHeight - 1, lastCol]].ClearContents();
                    }

                    ws.Cells[top, 1] = groups[i].Key;
                    int col = 1;
                    foreach (var s in groups[i].OrderBy(x => x.Start))
                    {
                        ws.Cells[top + 1, col] = s.Start.ToString("HH:mm") + " - " + s.End.ToString("HH:mm");
                        ws.Cells[top + 2, col] = "Зал " + s.Hall;
                        ws.Cells[top + 3, col] = s.Price.ToString("0.##", Ru) + " руб";
                        col++;
                        if (col > lastCol) break;
                    }
                }

                // если кинотеатров меньше 3 - лишние заготовки уже очищены
                if (File.Exists(path)) File.Delete(path);
                wb.SaveAs(path);
            }
            finally
            {
                if (wb != null) { wb.Close(false); ReleaseCom(wb); }
                app.Quit(); ReleaseCom(app);
            }
        }

        // =====================================================================
        //  3.4  сохранение расписания в xlsx ИЛИ html (формат - по фильтру диалога)
        // =====================================================================
        private void saveButton_Click(object sender, EventArgs e)
        {
            try
            {
                string path = AskSavePath(
                    "Excel (*.xlsx)|*.xlsx|HTML (*.html)|*.html", "Расписание");
                if (path == null) return;

                string ext = Path.GetExtension(path).ToLower();
                if (ext == ".html" || ext == ".htm")
                    SaveHtml(path, datePicker.Value);
                else
                {
                    if (ext != ".xlsx") path += ".xlsx";
                    SaveAllCinemasXlsx(path, datePicker.Value);
                }
                MessageBox.Show("Файл сохранён.");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void SaveHtml(string path, DateTime date)
        {
            var all = LoadSchedule(date, null, null);
            if (all.Count == 0) throw new InvalidOperationException("Сеансов на выбранный день нет.");

            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html><html><head><meta charset=\"utf-8\">");
            sb.AppendLine("<title>Расписание</title>");
            sb.AppendLine("<style>table{border-collapse:collapse;margin-bottom:20px}" +
                          "td,th{border:1px solid #444;padding:4px 10px}th{background:#ddd}</style>");
            sb.AppendLine("</head><body>");
            sb.AppendLine("<h1>Расписание на " + date.ToString("d MMMM yyyy", Ru) + "</h1>");

            foreach (var g in all.GroupBy(r => r.Cinema).OrderBy(g => g.Key))
            {
                sb.AppendLine("<h2>" + WebUtility(g.Key) + "</h2>");
                sb.AppendLine("<table><tr><th>Фильм</th><th>Начало сеанса</th><th>Зал</th><th>Цена</th></tr>");
                foreach (var r in g.OrderBy(x => x.Film).ThenBy(x => x.Start))
                    sb.AppendLine("<tr><td>" + WebUtility(r.Film) + "</td><td>" + r.Start.ToString("HH:mm") +
                                  "</td><td>" + r.Hall + "</td><td>" + r.Price.ToString("0.00", Ru) + " руб</td></tr>");
                sb.AppendLine("</table>");
            }
            sb.AppendLine("</body></html>");
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
        }

        private static string WebUtility(string s)
        {
            return System.Net.WebUtility.HtmlEncode(s);
        }

        private static void ShowError(Exception ex)
        {
            MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}

// =========================================================================
// Program.cs (если нужно): Application.Run(new LabWork22.Form1());
// =========================================================================

/*
ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ (для отчёта)

6.1 Какое пространство имен требуется подключить для работы с Word?
    using Word = Microsoft.Office.Interop.Word;
    (предварительно в проект добавляется ссылка на COM-библиотеку
    Microsoft Word Object Library; для Excel - Microsoft.Office.Interop.Excel).

6.2 Какие библиотеки позволяют сохранять данные в формате docx?
    - Microsoft.Office.Interop.Word (требует установленного MS Office);
    - Open XML SDK (DocumentFormat.OpenXml) - без установленного Office;
    - DocX (Xceed DocX / Novacode), NPOI (NPOI.XWPF), Aspose.Words, Spire.Doc.

6.3 Как выполнить экспорт данных в формате txt?
    Через класс File: File.WriteAllText / File.WriteAllLines (перезапись),
    File.AppendAllText / File.AppendAllLines (дозапись), либо через
    StreamWriter: using (var w = new StreamWriter(fileName)) { w.WriteLine(...); }.
    При необходимости последним параметром указывается кодировку (Encoding.UTF8).
*/
