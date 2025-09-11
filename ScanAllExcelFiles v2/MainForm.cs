using ClosedXML.Excel;
using ExcelDataReader;
using ExcelDataReader.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScannerWinForms
{
    public partial class MainForm : Form
    {
        private string lastResultFolder = "";

        public MainForm()
        {
            InitializeComponent();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            panelInput.Visible = true;
            panelProgress.Visible = false;
            panelResult.Visible = false;

            this.SizeChanged += (s, e) => UpdateLayout();
            UpdateLayout();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
                txtFolder.Text = dlg.SelectedPath;
        }

        private async void btnStart_Click(object sender, EventArgs e)
        {
            string rootFolder = txtFolder.Text.Trim();
            if (!Directory.Exists(rootFolder))
            {
                MessageBox.Show("Папка не найдена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var sheetNameFilter = txtSheet.Text.Trim();

            List<string> targetCells;
            try
            {
                targetCells = ParseCells(txtCells.Text.Trim());
                if (targetCells.Count == 0)
                    throw new Exception("Не указаны ячейки.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка в формате ячеек: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var excelFiles = Directory.GetFiles(rootFolder, "*.*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                         || f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
                         || f.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase)
                         || f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (excelFiles.Count == 0)
            {
                MessageBox.Show("Не найдено Excel файлов в выбранной папке.", "Инфо", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            panelInput.Visible = false;
            panelProgress.Visible = true;
            panelResult.Visible = false;
            lblStatus.Text = "Подготовка...";
            progressBar.Value = 0;
            progressBar.Maximum = Math.Max(1, excelFiles.Count);
            UpdateLayout();

            var progress = new Progress<int>(value =>
            {
                lblStatus.Text = $"Сканирование {value}/{excelFiles.Count}";
                progressBar.Value = Math.Clamp(value, progressBar.Minimum, progressBar.Maximum);
            });

            string resultFilePath = null;
            Exception backgroundException = null;

            await Task.Run(() =>
            {
                try
                {
                    var results = new List<List<object>>();
                    var errors = new List<List<object>>();

                    var header = new List<object> { "Файл (относительный путь)", "Уровень папки", "Имя файла", "Имя листа" };
                    header.AddRange(targetCells);
                    results.Add(header);
                    errors.Add(new List<object> { "Файл (относительный путь)", "Имя файла", "Описание ошибки" });

                    int current = 0;
                    foreach (var file in excelFiles)
                    {
                        current++;
                        (progress as IProgress<int>).Report(current);

                        if (Path.GetFileName(file).StartsWith("~$"))
                            continue;
                        try
                        {
                            bool sheetFound = ReadExcelFile(file, sheetNameFilter, targetCells, rootFolder, results);
                            if (!sheetFound)
                            {
                                errors.Add(new List<object>
                                {
                                    Path.GetRelativePath(rootFolder, file),
                                    Path.GetFileName(file),
                                    $"Не найден лист \"{sheetNameFilter}\""
                                });
                            }
                        }
                        catch (HeaderException hex)
                        {
                            errors.Add(new List<object>
                            {
                                Path.GetRelativePath(rootFolder, file),
                                Path.GetFileName(file),
                                "Файл зашифрован или повреждён"
                            });
                        }
                        catch (Exception ex)
                        {
                            errors.Add(new List<object>
                            {
                                Path.GetRelativePath(rootFolder, file),
                                Path.GetFileName(file),
                                ex.Message
                            });
                        }
                    }

                    string timestamp = DateTime.Now.ToString("yyyy.MM.dd HH-mm");
                    string resultFileName = $"{timestamp} Result.xlsx";

                    string exeDir = Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
                    resultFilePath = Path.Combine(exeDir, resultFileName);

                    using var workbook = new XLWorkbook();
                    var ws1 = workbook.Worksheets.Add("Собранные данные");
                    WriteSheet(ws1, results);

                    var ws2 = workbook.Worksheets.Add("Ошибки");
                    WriteSheet(ws2, errors);

                    workbook.SaveAs(resultFilePath);
                }
                catch (Exception ex)
                {
                    backgroundException = ex;
                }
            });

            if (backgroundException != null)
            {
                MessageBox.Show("Ошибка при сканировании: " + backgroundException.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                panelProgress.Visible = false;
                panelInput.Visible = true;
                UpdateLayout();
                return;
            }

            lastResultFolder = resultFilePath != null ? Path.GetDirectoryName(resultFilePath) : "";
            panelProgress.Visible = false;
            panelResult.Visible = true;
            lblFinished.Text = "✅ Сканирование завершено!";
            rtbFooter.Text = "\n\n(c) Галиев Ленар\nИсходный код: https://github.com/LEN4R/ScanAllExcelFiles_v2";
            UpdateLayout();
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(lastResultFolder) && Directory.Exists(lastResultFolder))
                System.Diagnostics.Process.Start("explorer.exe", lastResultFolder);
            else
                MessageBox.Show("Папка результата не найдена.", "Инфо", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnRestart_Click(object sender, EventArgs e)
        {
            txtFolder.Text = "";
            txtSheet.Text = "";
            txtCells.Text = "";
            panelResult.Visible = false;
            panelProgress.Visible = false;
            panelInput.Visible = true;
            UpdateLayout();
        }

        private void UpdateLayout()
        {
            var margin = 15;

            progressBar.Width = Math.Min(520, panelProgress.ClientSize.Width - margin * 2);
            progressBar.Left = (panelProgress.ClientSize.Width - progressBar.Width) / 2;
            lblStatus.Left = (panelProgress.ClientSize.Width - lblStatus.Width) / 2;

            int spacing = 12;
            int totalButtonsWidth = btnOpenFolder.Width + spacing + btnRestart.Width;
            int startX = (panelResult.ClientSize.Width - totalButtonsWidth) / 2;
            btnOpenFolder.Left = startX;
            btnRestart.Left = btnOpenFolder.Right + spacing;

            int baseHeight = panelInput.Visible ? panelInput.PreferredSize.Height : txtCells.Bottom + 40;
            int progressExtra = panelProgress.Visible ? progressBar.Height + lblStatus.Height + 50 : 0;
            int resultExtra = panelResult.Visible ? btnOpenFolder.Height + rtbFooter.Height + 60 : 0;

            int targetHeight = Math.Max(this.MinimumSize.Height, baseHeight + progressExtra + resultExtra);
            this.Height = targetHeight;
        }

        // === Парсинг ячеек ===
        private List<string> ParseCells(string input)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(input)) return result;

            var parts = input.Split(',', StringSplitOptions.RemoveEmptyEntries);
            var regexCell = new Regex(@"^[A-Za-z]+[0-9]+$", RegexOptions.Compiled);
            foreach (var raw in parts.Select(p => p.Trim()))
            {
                if (string.IsNullOrEmpty(raw)) continue;

                if (raw.Contains('-'))
                {
                    var rangeParts = raw.Split('-', StringSplitOptions.RemoveEmptyEntries);
                    if (rangeParts.Length != 2) throw new Exception($"Неверный диапазон: {raw}");
                    string start = rangeParts[0].Trim();
                    string end = rangeParts[1].Trim();

                    if (!regexCell.IsMatch(start) || !regexCell.IsMatch(end))
                        throw new Exception($"Диапазон должен содержать только английские буквы и цифры: {raw}");

                    string startCol = GetColumnLetters(start);
                    string endCol = GetColumnLetters(end);
                    int rowStart = GetRowNumber(start);
                    int rowEnd = GetRowNumber(end);
                    if (rowStart != rowEnd) throw new Exception($"В диапазоне строки должны совпадать: {raw}");

                    int colStart = ColLetterToIndex(startCol);
                    int colEnd = ColLetterToIndex(endCol);
                    if (colStart > colEnd) (colStart, colEnd) = (colEnd, colStart);

                    for (int c = colStart; c <= colEnd; c++)
                        result.Add(ColumnIndexToLetter(c) + rowStart);
                }
                else
                {
                    if (!regexCell.IsMatch(raw))
                        throw new Exception($"Ячейка должна содержать только английские буквы и цифры и начинаться с буквы: {raw}");
                    result.Add(raw);
                }
            }

            return result;
        }

        // === Универсальное чтение Excel через ExcelDataReader ===
        //private void ReadExcelFile(string file, string sheetNameFilter, List<string> targetCells, string rootFolder, List<List<object>> results)
        //{
        //    using var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        //    using var reader = ExcelReaderFactory.CreateReader(stream);

        //    do
        //    {
        //        string sheetName = reader.Name ?? $"Лист{Guid.NewGuid()}";
        //        if (!string.IsNullOrEmpty(sheetNameFilter) &&
        //            !sheetName.Equals(sheetNameFilter, StringComparison.OrdinalIgnoreCase))
        //            continue;

        //        var sheetData = new List<object[]>();
        //        while (reader.Read())
        //        {
        //            var values = new object[reader.FieldCount];
        //            reader.GetValues(values);
        //            sheetData.Add(values);
        //        }

        //        string relativePath = Path.GetRelativePath(rootFolder, file);
        //        string fileNameOnly = Path.GetFileNameWithoutExtension(file);
        //        int level = relativePath.Split(Path.DirectorySeparatorChar).Length - 1;
        //        var row = new List<object> { relativePath, level, fileNameOnly, sheetName };

        //        foreach (var addr in targetCells)
        //        {
        //            object value = "";
        //            try
        //            {
        //                int col = ColLetterToIndex(GetColumnLetters(addr));
        //                int r = GetRowNumber(addr) - 1;
        //                if (r >= 0 && r < sheetData.Count)
        //                {
        //                    var line = sheetData[r];
        //                    if (col >= 0 && col < line.Length)
        //                        value = line[col] ?? "";
        //                }
        //            }
        //            catch { }
        //            row.Add(value);
        //        }

        //        results.Add(row);

        //    } while (reader.NextResult());
        //}

        // === Универсальное чтение Excel через ExcelDataReader ===
        private bool ReadExcelFile(string file, string sheetNameFilter, List<string> targetCells, string rootFolder, List<List<object>> results)
        {
            bool sheetFound = false;

            using var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            do
            {
                string sheetName = reader.Name ?? $"Лист{Guid.NewGuid()}";
                if (!string.IsNullOrEmpty(sheetNameFilter) &&
                    !sheetName.Equals(sheetNameFilter, StringComparison.OrdinalIgnoreCase))
                    continue;

                sheetFound = true; // нашли нужный лист

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

                // Пробегаем все целевые ячейки
                var row = new List<object> { relativePath, level, fileNameOnly, sheetName };
                foreach (var addr in targetCells)
                {
                    object value = "";
                    try
                    {
                        int col = ColLetterToIndex(GetColumnLetters(addr));
                        int r = GetRowNumber(addr) - 1;
                        if (r >= 0 && r < sheetData.Count)
                        {
                            var line = sheetData[r];
                            if (col >= 0 && col < line.Length)
                                value = line[col] ?? "";
                        }
                    }
                    catch { }
                    row.Add(value);
                }

                // Если хотя бы одна ячейка не пустая — добавляем строку
                if (row.Skip(4).Any(v => v != null && v.ToString() != ""))
                    results.Add(row);

            } while (reader.NextResult());

            return sheetFound;
        }


        private void WriteSheet(IXLWorksheet ws, List<List<object>> data)
        {
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    object value = data[i][j];
                    if (value is double || value is int || value is decimal)
                        ws.Cell(i + 1, j + 1).Value = Convert.ToDouble(value);
                    else if (value is DateTime dt)
                        ws.Cell(i + 1, j + 1).Value = dt;
                    else
                        ws.Cell(i + 1, j + 1).Value = value?.ToString() ?? "";
                }
            }

            ws.Row(1).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();
        }

        // === Вспомогательные ===
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
