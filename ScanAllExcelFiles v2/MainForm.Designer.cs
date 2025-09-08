namespace ExcelScannerWinForms
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtSheet;
        private System.Windows.Forms.TextBox txtCells;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.Label lblSheet;
        private System.Windows.Forms.Label lblCells;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblResultPath;
        private System.Windows.Forms.Button btnOpenFolder;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtSheet = new System.Windows.Forms.TextBox();
            this.txtCells = new System.Windows.Forms.TextBox();
            this.btnStart = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblFolder = new System.Windows.Forms.Label();
            this.lblSheet = new System.Windows.Forms.Label();
            this.lblCells = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblResultPath = new System.Windows.Forms.Label();
            this.btnOpenFolder = new System.Windows.Forms.Button();
            this.SuspendLayout();

            int margin = 15;
            int btnHeight = 27; // фиксированная высота кнопки «Обзор»
            int labelSpacing = 5;

            // ---------- Папка ----------
            this.lblFolder.Text = "Выберите папку с Excel-файлами:";
            this.lblFolder.Location = new System.Drawing.Point(margin, 10);
            this.lblFolder.AutoSize = true;

            this.txtFolder.Location = new System.Drawing.Point(margin, lblFolder.Bottom + labelSpacing);
            this.txtFolder.Width = 500;
            this.txtFolder.Height = btnHeight; // выравниваем по высоте с кнопкой

            this.btnBrowse.Text = "Обзор...";
            this.btnBrowse.Location = new System.Drawing.Point(txtFolder.Right + 10, txtFolder.Top);
            this.btnBrowse.Width = 100;
            this.btnBrowse.Height = btnHeight; // фиксированная высота
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);

            // ---------- Лист ----------
            this.lblSheet.Text = "Введите название листа (оставьте пустым для всех):";
            this.lblSheet.Location = new System.Drawing.Point(margin, txtFolder.Bottom + 10);
            this.lblSheet.AutoSize = true;

            this.txtSheet.Location = new System.Drawing.Point(margin, lblSheet.Bottom + labelSpacing);
            this.txtSheet.Width = txtFolder.Width + btnBrowse.Width + 10; // чтобы совпадала ширина
            this.txtSheet.Height = 65; // выравниваем по высоте кнопки

            // ---------- Ячейки ----------
            this.lblCells.Text = "Введите ячейки (например: D9,AA9-BE9):";
            this.lblCells.Location = new System.Drawing.Point(margin, txtSheet.Bottom + 10);
            this.lblCells.AutoSize = true;

            this.txtCells.Location = new System.Drawing.Point(margin, lblCells.Bottom + labelSpacing);
            this.txtCells.Width = txtSheet.Width;
            this.txtCells.Height = btnHeight; // выравниваем

            // ---------- Кнопка старт ----------
            this.btnStart.Text = "Сканировать";
            this.btnStart.Location = new System.Drawing.Point(margin, txtCells.Bottom + 20);
            this.btnStart.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            this.btnStart.Width = 50;
            this.btnStart.Height = btnHeight;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);

            // ---------- Статус и прогрессбар ----------
            this.lblStatus.Text = "";
            this.lblStatus.Location = new System.Drawing.Point(margin, btnStart.Bottom + 10);
            this.lblStatus.AutoSize = true;
            this.lblStatus.Visible = false;

            this.progressBar.Location = new System.Drawing.Point(margin, lblStatus.Bottom + 5);
            this.progressBar.Width = txtSheet.Width;
            this.progressBar.Height = btnHeight;
            this.progressBar.Visible = false;

            // ---------- Результат ----------
            this.lblResultPath.Text = "";
            this.lblResultPath.Location = new System.Drawing.Point(margin, progressBar.Bottom + 10);
            this.lblResultPath.AutoSize = true;
            this.lblResultPath.Visible = false;

            this.btnOpenFolder.Text = "Открыть папку";
            this.btnOpenFolder.Location = new System.Drawing.Point(margin, lblResultPath.Bottom + 5);
            this.btnOpenFolder.Width = 150;
            this.btnOpenFolder.Height = btnHeight;
            this.btnOpenFolder.Visible = false;
            this.btnOpenFolder.Click += new System.EventHandler(this.btnOpenFolder_Click);

            // ---------- Форма ----------
            this.ClientSize = new System.Drawing.Size(txtSheet.Width + margin * 2, btnStart.Bottom + 20 + btnHeight * 1);
            this.Controls.Add(this.lblFolder);
            this.Controls.Add(this.txtFolder);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.lblSheet);
            this.Controls.Add(this.txtSheet);
            this.Controls.Add(this.lblCells);
            this.Controls.Add(this.txtCells);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblResultPath);
            this.Controls.Add(this.btnOpenFolder);
            this.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.Text = "ScanAllExcelFiles v2";
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
