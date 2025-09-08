using ClosedXML.Excel;
using ExcelDataReader;
using ExcelDataReader.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScannerWinForms
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // начальный размер окна
            this.MinimumSize = new System.Drawing.Size(710, 280);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            progressBar.Visible = false;
            lblStatus.Visible = false;
            lblResultPath.Visible = false;
            btnOpenFolder.Visible = false;

            this.Resize += MainForm_Resize;
            UpdateFormHeight();
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            // Растягиваем кнопки и прогресс-бар по ширине окна
            int margin = 15;
            btnStart.Width = this.ClientSize.Width - 2 * margin;
            progressBar.Width = this.ClientSize.Width - 2 * margin;
            lblResultPath.Width = this.ClientSize.Width - 2 * margin;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                    txtFolder.Text = dlg.SelectedPath;
            }
        }

        private async void btnStart_Click(object sender, EventArgs e)
        {
            string rootFolder = txtFolder.Text;
            if (!Directory.Exists(rootFolder))
            {
                MessageBox.Show("Папка не найдена!");
                return;
            }

            var sheetNameFilter = txtSheet.Text.Trim();

            // Обработка ячеек
            List<string> targetCells;
            try
            {
                targetCells = ParseCells(txtCells.Text.Trim());
            }
            catch
            {
                MessageBox.Show("Ошибка в формате ячеек.");
                return;
            }

            // Подготовка интерфейса
            btnStart.Visible = false;
            lblStatus.Visible = true;
            progressBar.Visible = true;
            lblResultPath.Visible = false;
            btnOpenFolder.Visible = false;
            lblStatus.Text = "Сканирование...";
            UpdateFormHeight();

            var results = new List<List<string>>();
            var errors = new List<List<string>>();

            // Заголовки
            var header = new List<string> { "Файл (относительный путь)", "Уровень папки", "Имя файла", "Имя листа" };
            header.AddRange(targetCells);
            results.Add(header);
            errors.Add(new List<string> { "Файл (относительный путь)", "Имя файла", "Описание ошибки" });

            // Получаем файлы
            var excelFiles = Directory.GetFiles(rootFolder, "*.*", SearchOption.AllDirectories)
                                      .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                                               || f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
                                               || f.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase))
                                      .ToList();

            int total = excelFiles.Count;
            int current = 0;

            foreach (var file in excelFiles)
            {
                current++;
                lblStatus.Text = $"Сканирование {current}/{total}";
                progressBar.Maximum = total;
                progressBar.Value = current;
                Application.DoEvents(); // обновляем UI

                if (Path.GetFileName(file).StartsWith("~$"))
                    continue;

                try
                {
                    if (file.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase))
                        ReadXlsbFile(file, sheetNameFilter, targetCells, rootFolder, results);
                    else
                        ReadXlsxFile(file, sheetNameFilter, targetCells, rootFolder, results);
                }
                catch (HeaderException hex)
                {
                    errors.Add(new List<string>
                    {
                        Path.GetRelativePath(rootFolder, file),
                        Path.GetFileName(file),
                        $"Файл зашифрован или повреждён: {hex.Message}"
                    });
                }
                catch (Exception ex)
                {
                    errors.Add(new List<string>
                    {
                        Path.GetRelativePath(rootFolder, file),
                        Path.GetFileName(file),
                        ex.Message
                    });
                }

                await Task.Delay(1); // для плавности UI
            }

            // Сохраняем результат
            string timestamp = DateTime.Now.ToString("yyyy.MM.dd HH-mm");
            string resultFileName = $"{timestamp} Result.xlsx";
            string resultFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, resultFileName);

            using (var workbook = new XLWorkbook())
            {
                var ws1 = workbook.Worksheets.Add("Собранные данные");
                WriteSheet(ws1, results);

                var ws2 = workbook.Worksheets.Add("Ошибки");
                WriteSheet(ws2, errors);

                workbook.SaveAs(resultFilePath);
            }

            // Завершение сканирования
            progressBar.Visible = false;
            lblStatus.Text = "✅ Сканирование выполнено!";
            lblResultPath.Text = resultFilePath;
            lblResultPath.Visible = true;
            btnOpenFolder.Visible = true;

            // Расположение результата
            int margin = 15;
            lblResultPath.Top = progressBar.Visible ? progressBar.Bottom + 10 : txtCells.Bottom + 10;
            lblResultPath.Left = margin;
            btnOpenFolder.Top = lblResultPath.Bottom + 5;
            btnOpenFolder.Left = margin;

            UpdateFormHeight();
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            string folder = Path.GetDirectoryName(lblResultPath.Text);
            if (Directory.Exists(folder))
                System.Diagnostics.Process.Start("explorer.exe", folder);
        }

        // ===================== Динамический размер формы =====================
        private void UpdateFormHeight()
        {
            int baseHeight = txtCells.Bottom + 20; // после полей ввода
            int progressHeight = progressBar.Visible ? progressBar.Height + 10 : 0;
            int resultHeight = (lblResultPath.Visible || btnOpenFolder.Visible) ? lblResultPath.Height + btnOpenFolder.Height + 20 : 0;
            this.Height = baseHeight + progressHeight + resultHeight;
        }

        // ===================== Парсинг ячеек =====================
        private List<string> ParseCells(string input)
        {
            var result = new List<string>();
            var parts = input.Split(',', StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                if (p.Contains("-"))
                {
                    var range = p.Split('-');
                    string start = range[0].Trim();
                    string end = range[1].Trim();
                    int row = GetRowNumber(start);
                    int startCol = ColLetterToIndex(GetColumnLetters(start));
                    int endCol = ColLetterToIndex(GetColumnLetters(end));

                    for (int c = startCol; c <= endCol; c++)
                        result.Add(ColumnIndexToLetter(c) + row);
                }
                else
                {
                    result.Add(p.Trim());
                }
            }
            return result;
        }

        // ===================== Чтение Excel =====================
        private void ReadXlsxFile(string file, string sheetNameFilter, List<string> targetCells, string rootFolder, List<List<string>> results)
        {
            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var sheetsToScan = string.IsNullOrEmpty(sheetNameFilter)
                    ? workbook.Worksheets.ToList()
                    : workbook.Worksheets
                              .Where(s => s.Name.Equals(sheetNameFilter, StringComparison.OrdinalIgnoreCase))
                              .ToList();

                if (!sheetsToScan.Any())
                    throw new Exception($"Лист '{sheetNameFilter}' не найден");

                foreach (var ws in sheetsToScan)
                {
                    string relativePath = Path.GetRelativePath(rootFolder, file);
                    string fileNameOnly = Path.GetFileNameWithoutExtension(file);
                    int level = relativePath.Split(Path.DirectorySeparatorChar).Length - 1;
                    var row = new List<string> { relativePath, level.ToString(), fileNameOnly, ws.Name };

                    foreach (var addr in targetCells)
                    {
                        string value;
                        try { value = ws.Cell(addr).GetFormattedString(); }
                        catch { value = ""; }
                        row.Add(value);
                    }
                    results.Add(row);
                }
            }
        }

        private void ReadXlsbFile(string file, string sheetNameFilter, List<string> targetCells, string rootFolder, List<List<string>> results)
        {
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    string sheetName = reader.Name ?? $"Лист{Guid.NewGuid()}";
                    if (!string.IsNullOrEmpty(sheetNameFilter) &&
                        !sheetName.Equals(sheetNameFilter, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var sheetData = new List<object[]>();
                    while (reader.Read())
                    {
                        var values = new object[reader.FieldCount];
                        reader.GetValues(values);
                        sheetData.Add(values);
                    }

                    string relativePath = Path.GetRelativePath(rootFolder, file);
                    string fileNameOnly = Path.GetFileNameWithoutExtension(file);
                    int level = relativePath.Split(Path.DirectorySeparatorChar).Length - 1;
                    var row = new List<string> { relativePath, level.ToString(), fileNameOnly, sheetName };

                    foreach (var addr in targetCells)
                    {
                        string value = "";
                        try
                        {
                            int col = ColLetterToIndex(GetColumnLetters(addr));
                            int r = GetRowNumber(addr) - 1;
                            if (r >= 0 && r < sheetData.Count)
                            {
                                var line = sheetData[r];
                                if (col >= 0 && col < line.Length)
                                    value = line[col]?.ToString() ?? "";
                            }
                        }
                        catch { }
                        row.Add(value);
                    }

                    results.Add(row);

                } while (reader.NextResult());
            }
        }

        // ===================== Сохранение в Excel =====================
        private void WriteSheet(IXLWorksheet ws, List<List<string>> data)
        {
            for (int i = 0; i < data.Count; i++)
                for (int j = 0; j < data[i].Count; j++)
                    ws.Cell(i + 1, j + 1).Value = data[i][j];

            ws.Row(1).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();
        }

        // ===================== Вспомогательные методы =====================
        private string GetColumnLetters(string cellAddress) => new string(cellAddress.Where(char.IsLetter).ToArray());

        private int GetRowNumber(string cellAddress)
        {
            var num = new string(cellAddress.Where(char.IsDigit).ToArray());
            return int.TryParse(num, out int r) ? r : -1;
        }

        private int ColLetterToIndex(string colLetters)
        {
            int sum = 0;
            foreach (char c in colLetters.ToUpperInvariant())
                sum = checked(sum * 26 + (c - 'A' + 1));
            return sum - 1;
        }

        private string ColumnIndexToLetter(int col)
        {
            string result = "";
            col++;
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                result = (char)(rem + 'A') + result;
                col = (col - 1) / 26;
            }
            return result;
        }
    }
}
