namespace ExcelScannerWinForms
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        // панели
        private System.Windows.Forms.Panel panelInput;
        private System.Windows.Forms.Panel panelProgress;
        private System.Windows.Forms.Panel panelResult;

        // input controls
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtSheet;
        private System.Windows.Forms.TextBox txtCells;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.Label lblSheet;
        private System.Windows.Forms.Label lblCells;

        // progress controls
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblStatus;

        // result controls
        private System.Windows.Forms.Label lblFinished;
        private System.Windows.Forms.Button btnOpenFolder;
        private System.Windows.Forms.Button btnRestart;
        private System.Windows.Forms.RichTextBox rtbFooter;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.panelInput = new System.Windows.Forms.Panel();
            this.lblFolder = new System.Windows.Forms.Label();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lblSheet = new System.Windows.Forms.Label();
            this.txtSheet = new System.Windows.Forms.TextBox();
            this.lblCells = new System.Windows.Forms.Label();
            this.txtCells = new System.Windows.Forms.TextBox();
            this.btnStart = new System.Windows.Forms.Button();

            this.panelProgress = new System.Windows.Forms.Panel();
            this.lblStatus = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();

            this.panelResult = new System.Windows.Forms.Panel();
            this.lblFinished = new System.Windows.Forms.Label();
            this.btnOpenFolder = new System.Windows.Forms.Button();
            this.btnRestart = new System.Windows.Forms.Button();
            this.rtbFooter = new System.Windows.Forms.RichTextBox();

            this.SuspendLayout();

            int margin = 15;
            int ctrlHeight = 35;
            int formWidth = 675;

            // ---- panelInput ----
            this.panelInput.Dock = System.Windows.Forms.DockStyle.Fill;

            // Папка
            this.lblFolder.Text = "Выберите папку с Excel-файлами:";
            this.lblFolder.AutoSize = true;
            this.lblFolder.Location = new System.Drawing.Point(margin, margin);

            this.txtFolder.Location = new System.Drawing.Point(margin, this.lblFolder.Bottom + 5);
            this.txtFolder.Width = 500;
            this.txtFolder.Height = ctrlHeight;

            this.btnBrowse.Text = "Обзор...";
            this.btnBrowse.Location = new System.Drawing.Point(this.txtFolder.Right + 8, this.txtFolder.Top);
            this.btnBrowse.Width = 120;
            this.btnBrowse.Height = ctrlHeight-8;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);

            // Лист
            this.lblSheet.Text = "Введите название листа (оставьте пустым для всех):";
            this.lblSheet.AutoSize = true;
            this.lblSheet.Location = new System.Drawing.Point(margin, this.txtFolder.Bottom + 10);

            this.txtSheet.Location = new System.Drawing.Point(margin, this.lblSheet.Bottom + 5);
            this.txtSheet.Width = 627; // фиксированная ширина
            this.txtSheet.Height = ctrlHeight;

            // Ячейки
            this.lblCells.Text = "Введите ячейки (например: D9,AA9-BE9):";
            this.lblCells.AutoSize = true;
            this.lblCells.Location = new System.Drawing.Point(margin, this.txtSheet.Bottom + 10);

            this.txtCells.Location = new System.Drawing.Point(margin, this.lblCells.Bottom + 5);
            this.txtCells.Width = 627; // фиксированная ширина
            this.txtCells.Height = ctrlHeight;

            // Кнопка Сканировать
            this.btnStart.Text = "Сканировать";
            this.btnStart.Width = 627; // фиксированная ширина
            this.btnStart.Height = ctrlHeight;
            this.btnStart.Location = new System.Drawing.Point(margin, this.txtCells.Bottom + 25);
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);

            this.panelInput.Controls.AddRange(new System.Windows.Forms.Control[]
            {
            this.lblFolder, this.txtFolder, this.btnBrowse,
            this.lblSheet, this.txtSheet,
            this.lblCells, this.txtCells,
            this.btnStart
            });

            // ---- panelProgress ----
            this.panelProgress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelProgress.Visible = false;

            this.lblStatus.Text = "Сканирование...";
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(margin, 70);

            this.progressBar.Location = new System.Drawing.Point(margin, this.lblStatus.Bottom + 5);
            this.progressBar.Width = 627;
            this.progressBar.Height = ctrlHeight;

            this.panelProgress.Controls.AddRange(new System.Windows.Forms.Control[]
                {
                    this.lblStatus, this.progressBar
                });

            // ---- panelResult ----
            this.panelResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelResult.Visible = false;

            this.lblFinished.Text = "✅ Сканирование завершено!";
            this.lblFinished.AutoSize = true;
            this.lblFinished.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold);
            this.lblFinished.Location = new System.Drawing.Point(margin, 50);

            this.btnOpenFolder.Text = "Открыть папку с результатом";
            this.btnOpenFolder.Width = 305;
            this.btnOpenFolder.Height = ctrlHeight;
            this.btnOpenFolder.Location = new System.Drawing.Point(margin, this.lblFinished.Bottom + 18);
            this.btnOpenFolder.Click += new System.EventHandler(this.btnOpenFolder_Click);

            this.btnRestart.Text = "Начать заново";
            this.btnRestart.Width = 305;
            this.btnRestart.Height = ctrlHeight;
            this.btnRestart.Location = new System.Drawing.Point(this.btnOpenFolder.Right + 12, this.lblFinished.Bottom + 18);
            this.btnRestart.Click += new System.EventHandler(this.btnRestart_Click);

            // Футер
            this.rtbFooter.ReadOnly = true;                                                                                     // Делает RichTextBox только для чтения, чтобы пользователь мог выделять и копировать текст, 
            this.rtbFooter.BackColor = System.Drawing.SystemColors.Control;                                                     // Установка фона (можно удалить, так как уже задано выше).
            this.rtbFooter.BorderStyle = System.Windows.Forms.BorderStyle.None;                                                 // Убирает границу вокруг RichTextBox, чтобы он выглядел как часть формы без рамки.
            this.rtbFooter.DetectUrls = true;                                                                                   // Включает автоматическое распознавание URL-адресов в тексте, делая их кликабельными ссылками.
            this.rtbFooter.Location = new System.Drawing.Point(margin, this.btnRestart.Bottom + 40);                            // Задаёт позицию RichTextBox на форме: X = отступ слева, Y = чуть ниже кнопки btnRestart (+5 пикселей).
            this.rtbFooter.Width = formWidth - 2 * margin;                                                                      // Задаёт ширину RichTextBox, равную ширине формы минус двойной отступ слева и справа.
            this.rtbFooter.Height = 70;                                                                                         // Задаёт фиксированную высоту RichTextBox в 70 пикселей.
            this.rtbFooter.Text = "(c) Галиев Ленар | Исходный код: https://github.com/LEN4R/ScanAllExcelFiles_v2";             
            this.rtbFooter.TabStop = false;                                                                                     // Делает так, чтобы фокус с клавиатуры (Tab) не переходил на этот RichTextBox.

            this.panelResult.Controls.AddRange(new System.Windows.Forms.Control[]
                {
                    this.lblFinished, this.btnOpenFolder, this.btnRestart, this.rtbFooter
                });


            // ---- Форма ----
            this.AutoSize = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.ClientSize = new System.Drawing.Size(formWidth, 300);
            this.MinimumSize = new System.Drawing.Size(formWidth, 300);
            this.MaximumSize = new System.Drawing.Size(formWidth, 300);
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                this.panelInput, this.panelProgress, this.panelResult
            });

            this.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.Text = "ScanAllExcelFiles v2";

            this.ResumeLayout(false);
        }

    }
}
