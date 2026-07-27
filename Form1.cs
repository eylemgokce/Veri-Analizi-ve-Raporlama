using Microsoft.VisualBasic.ApplicationServices;
using OfficeOpenXml.Utils;
using Raporlama.Models;
using Raporlama.Services;
using System.ComponentModel;
using System.Drawing.Drawing2D;
using System.Security.Cryptography.Xml;

namespace Raporlama
{
    public partial class frmMain : Form
    {
        private List<ExcelDataRow> _kayitlar = new();
        private string? _sonRaporYolu;

        // Ana kontroller
        private DataGridView _dgvData = null!;
        private Button _btnSelectExcel = null!;
        private Button _btnAnalyze = null!;
        private Button _btnExportReport = null!;
        private Button _btnSendMail = null!;

        // Sayaçlar
        private Label _lblTotalValue = null!;
        private Label _lblSuccessValue = null!;
        private Label _lblErrorValue = null!;

        // Dosya / durum
        private Label _lblFileName = null!;
        private Label _lblStatus = null!;

        // Hata listesi
        private ListBox _lstErrors = null!;

        // Grafik
        private DonutChart _donutChart = null!;

        public frmMain()
        {
            // Designer'daki eski ekranı kullanmıyoruz.
            // Arayüzü tamamen kodla oluşturuyoruz.
            BuildDashboard();
        }

        private void btnSelectExcel_Click(object sender, EventArgs e)
        {
            btnSelectExcel_Click(sender, e);
        }

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            btnAnalyze_Click(sender, e);
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
        }

        private void frmMain_Load(object sender, EventArgs e)
        { }
        private void BuildDashboard()
        {
            // ---------------- FORM ----------------

            Text = "Üretim Raporlama ve Veri Doğrulama Sistemi";
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(1400, 820);
            MinimumSize = new Size(1200, 700);

            BackColor = Color.FromArgb(244, 246, 251);
            Font = new Font("Segoe UI", 9F);

            // ---------------- SIDEBAR ----------------

            Panel sidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = 75,
                BackColor = Color.FromArgb(78, 82, 200)
            };

            Controls.Add(sidebar);

            Label logo = new Label
            {
                Text = "R",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 75
            };

            sidebar.Controls.Add(logo);

            AddSideMenuButton(sidebar, "⌂", 100);
            AddSideMenuButton(sidebar, "▦", 160);
            AddSideMenuButton(sidebar, "✓", 220);
            AddSideMenuButton(sidebar, "⚙", 280);

            // ---------------- HEADER ----------------

            Panel header = new Panel
            {
                Location = new Point(75, 0),
                Height = 85,
                Width = ClientSize.Width - 75,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.White
            };

            Controls.Add(header);

            Label title = new Label
            {
                Text = "ÜRETİM RAPORLAMA",
                Location = new Point(30, 18),
                AutoSize = true,
                ForeColor = Color.FromArgb(38, 42, 66),
                Font = new Font("Segoe UI", 17, FontStyle.Bold)
            };

            Label subtitle = new Label
            {
                Text = "Üretim verilerini yükleyin, doğrulayın ve analiz edin",
                Location = new Point(32, 52),
                AutoSize = true,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 9)
            };

            header.Controls.Add(title);
            header.Controls.Add(subtitle);

            // ---------------- İÇERİK ----------------

            Panel content = new Panel
            {
                Location = new Point(100, 105),
                Size = new Size(ClientSize.Width - 125, ClientSize.Height - 130),
                Anchor = AnchorStyles.Top |
                         AnchorStyles.Bottom |
                         AnchorStyles.Left |
                         AnchorStyles.Right,
                BackColor = Color.Transparent
            };

            Controls.Add(content);

            // ---------------- BUTONLAR ----------------

            _btnSelectExcel = CreateButton(
                "Excel Dosyası Seç",
                Color.FromArgb(72, 105, 246));

            _btnSelectExcel.Location = new Point(0, 0);
            _btnSelectExcel.Click += BtnSelectExcel_Click;

            _btnAnalyze = CreateButton(
                "Analiz Et",
                Color.FromArgb(24, 190, 145));

            _btnAnalyze.Location = new Point(190, 0);
            _btnAnalyze.Click += BtnAnalyze_Click;

            content.Controls.Add(_btnSelectExcel);
            content.Controls.Add(_btnAnalyze);

            _btnExportReport = CreateButton(
            "Rapor Oluştur",
             Color.FromArgb(245, 158, 11));

            _btnExportReport.Location = new Point(380, 0);
            _btnExportReport.Click += BtnExportReport_Click;

            content.Controls.Add(_btnExportReport);

            _btnSendMail = CreateButton(
            "Mail Gönder",
             Color.FromArgb(124, 58, 237));

            _btnSendMail.Location = new Point(570, 0);
            _btnSendMail.Click += BtnSendMail_Click;

            content.Controls.Add(_btnSendMail);

            _lblFileName = new Label
            {
                Text = "Henüz dosya seçilmedi",
                Location = new Point(765, 12),
                AutoSize = true,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 9)
            };

            content.Controls.Add(_lblFileName);

            // ---------------- KPI KARTLARI ----------------

            Panel totalCard = CreateCard(
                "TOPLAM KAYIT",
                "0",
                Color.FromArgb(82, 114, 255),
                out _lblTotalValue);

            Panel successCard = CreateCard(
                "BAŞARILI",
                "0",
                Color.FromArgb(29, 201, 151),
                out _lblSuccessValue);

            Panel errorCard = CreateCard(
                "HATALI",
                "0",
                Color.FromArgb(247, 89, 129),
                out _lblErrorValue);

            totalCard.Location = new Point(0, 65);
            successCard.Location = new Point(270, 65);
            errorCard.Location = new Point(540, 65);

            content.Controls.Add(totalCard);
            content.Controls.Add(successCard);
            content.Controls.Add(errorCard);

            // ---------------- TABLO PANELİ ----------------

            Panel tableCard = CreateWhitePanel();

            tableCard.Location = new Point(0, 190);
            tableCard.Size = new Size(
                content.Width - 365,
                content.Height - 240);

            tableCard.Anchor = AnchorStyles.Top |
                               AnchorStyles.Bottom |
                               AnchorStyles.Left |
                               AnchorStyles.Right;

            Label tableTitle = new Label
            {
                Text = "Üretim Kayıtları",
                Location = new Point(20, 15),
                AutoSize = true,
                ForeColor = Color.FromArgb(45, 49, 70),
                Font = new Font("Segoe UI", 11, FontStyle.Bold)
            };

            tableCard.Controls.Add(tableTitle);

            _dgvData = new DataGridView
            {
                Location = new Point(20, 50),
                Size = new Size(
                    tableCard.Width - 40,
                    tableCard.Height - 70),

                Anchor = AnchorStyles.Top |
                         AnchorStyles.Bottom |
                         AnchorStyles.Left |
                         AnchorStyles.Right,

                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,

                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,

                AutoSizeColumnsMode =
                    DataGridViewAutoSizeColumnsMode.Fill,

                SelectionMode =
                    DataGridViewSelectionMode.FullRowSelect,

                MultiSelect = false
            };

            StyleDataGridView();

            tableCard.Controls.Add(_dgvData);
            content.Controls.Add(tableCard);

            // ---------------- SAĞ TARAF ----------------

            Panel rightPanel = new Panel
            {
                Location = new Point(content.Width - 340, 190),
                Size = new Size(340, content.Height - 240),

                Anchor = AnchorStyles.Top |
                         AnchorStyles.Bottom |
                         AnchorStyles.Right,

                BackColor = Color.Transparent
            };

            content.Controls.Add(rightPanel);

            // ---------------- DONUT ----------------

            Panel chartCard = CreateWhitePanel();

            chartCard.Location = new Point(0, 0);
            chartCard.Size = new Size(340, 280);

            Label chartTitle = new Label
            {
                Text = "Doğrulama Sonucu",
                Location = new Point(20, 15),
                AutoSize = true,
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(45, 49, 70)
            };

            chartCard.Controls.Add(chartTitle);

            _donutChart = new DonutChart
            {
                Location = new Point(65, 55),
                Size = new Size(210, 190),
                SuccessPercentage = 0
            };

            chartCard.Controls.Add(_donutChart);
            rightPanel.Controls.Add(chartCard);

            // ---------------- HATA LİSTESİ ----------------

            Panel errorListCard = CreateWhitePanel();

            errorListCard.Location = new Point(0, 295);
            errorListCard.Size = new Size(
                340,
                Math.Max(180, rightPanel.Height - 295));

            errorListCard.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Bottom |
                AnchorStyles.Left |
                AnchorStyles.Right;

            Label errorTitle = new Label
            {
                Text = "Hata Detayları",
                Location = new Point(20, 15),
                AutoSize = true,
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(45, 49, 70)
            };

            _lstErrors = new ListBox
            {
                Location = new Point(20, 48),
                Size = new Size(
                    errorListCard.Width - 40,
                    errorListCard.Height - 70),

                Anchor = AnchorStyles.Top |
                         AnchorStyles.Bottom |
                         AnchorStyles.Left |
                         AnchorStyles.Right,

                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9),
                BackColor = Color.White
            };

            errorListCard.Controls.Add(errorTitle);
            errorListCard.Controls.Add(_lstErrors);

            rightPanel.Controls.Add(errorListCard);

            // ---------------- STATUS ----------------

            _lblStatus = new Label
            {
                Text = "● Sistem hazır",
                AutoSize = true,
                ForeColor = Color.FromArgb(29, 201, 151),

                Location = new Point(
                    100,
                    ClientSize.Height - 30),

                Anchor = AnchorStyles.Bottom |
                         AnchorStyles.Left
            };

            Controls.Add(_lblStatus);
        }

        // ----------------------------------------------------
        // EXCEL DOSYASI SEÇ
        // ----------------------------------------------------

        private void BtnSelectExcel_Click(object? sender, EventArgs e)
        {
            using OpenFileDialog dialog = new();

            dialog.Filter =
                "Excel Files (*.xlsx)|*.xlsx";

            if (dialog.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                ExcelService excelService = new();

                _kayitlar =
                excelService.ReadDynamicExcel(dialog.FileName);

                ShowDynamicData();

                _lblTotalValue.Text =
                    _kayitlar.Count.ToString();

                _lblSuccessValue.Text = "-";
                _lblErrorValue.Text = "-";

                _lblFileName.Text =
                    Path.GetFileName(dialog.FileName);

                _lblStatus.Text =
                    "● Excel dosyası başarıyla yüklendi";

                _lblStatus.ForeColor =
                    Color.FromArgb(29, 201, 151);

                _lstErrors.Items.Clear();

                _donutChart.SuccessPercentage = 0;
                _donutChart.Invalidate();

                

               ResetRowColors();
               
    }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Dosya Okuma Hatası",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                _lblStatus.Text =
                    "● Dosya yüklenemedi";

                _lblStatus.ForeColor =
                    Color.FromArgb(247, 89, 129);
            }
        }

        // ----------------------------------------------------
        // ANALİZ
        // ----------------------------------------------------

        private void BtnAnalyze_Click(object? sender, EventArgs e)
        {
            if (_kayitlar.Count == 0)
            {
                MessageBox.Show(
                    "Önce bir Excel dosyası seçmelisiniz.",
                    "Uyarı",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            ValidationService validationService = new();

            validationService.ValidateDynamic(_kayitlar);

            int basarili = 0;
            int hatali = 0;

            _lstErrors.Items.Clear();

            for (int i = 0; i < _kayitlar.Count; i++)
            {
                ExcelDataRow row = _kayitlar[i];

                if (row.IsValid)
                {
                    basarili++;

                    _dgvData.Rows[i].DefaultCellStyle.BackColor =
                        Color.FromArgb(229, 250, 242);
                }
                else
                {
                    hatali++;

                    _dgvData.Rows[i].DefaultCellStyle.BackColor =
                        Color.FromArgb(255, 232, 238);

                    foreach (string hata in row.Errors)
                    {
                        _lstErrors.Items.Add(
                            $"Satır {row.RowNumber}: {hata}");
                    }
                }
            }

            _lblTotalValue.Text = _kayitlar.Count.ToString();
            _lblSuccessValue.Text = basarili.ToString();
            _lblErrorValue.Text = hatali.ToString();

            double basariYuzdesi =
                _kayitlar.Count == 0
                    ? 0
                    : (double)basarili / _kayitlar.Count * 100;

            _donutChart.SuccessPercentage = basariYuzdesi;
            _donutChart.Invalidate();

            _lblStatus.Text =
                $"● Analiz tamamlandı - {hatali} hatalı kayıt bulundu";

            _lblStatus.ForeColor =
                hatali == 0
                    ? Color.FromArgb(29, 201, 151)
                    : Color.FromArgb(247, 89, 129);
        }

        private void BtnExportReport_Click(object? sender, EventArgs e)
        {
            if (_kayitlar.Count == 0)
            {
                MessageBox.Show(
                    "Önce bir Excel dosyası seçmelisiniz.",
                    "Uyarı",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            using SaveFileDialog dialog = new();

            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            string kaynakDosyaAdi =
            Path.GetFileNameWithoutExtension(_lblFileName.Text);

            dialog.FileName =
                $"{kaynakDosyaAdi}_Analiz_Raporu_{DateTime.Now:ddMMyyyy_HHmm}.xlsx";

            if (dialog.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                ExcelService excelService = new();

                excelService.CreateDynamicReport(
                    dialog.FileName,
                    _kayitlar);
                _sonRaporYolu = dialog.FileName;

                _lblStatus.Text = "● Rapor başarıyla oluşturuldu";
                _lblStatus.ForeColor = Color.FromArgb(29, 201, 151);

                MessageBox.Show(
                    "Rapor başarıyla oluşturuldu.",
                    "Başarılı",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "● Rapor oluşturulamadı";
                _lblStatus.ForeColor = Color.FromArgb(247, 89, 129);

                MessageBox.Show(
                    ex.Message,
                    "Rapor Oluşturma Hatası",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void BtnSendMail_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_sonRaporYolu) ||
                !File.Exists(_sonRaporYolu))
            {
                MessageBox.Show(
                    "Önce bir rapor oluşturmalısınız.",
                    "Uyarı",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            using MailForm mailForm = new(_sonRaporYolu);
            mailForm.ShowDialog();
        }

        // ----------------------------------------------------
        // DATAGRIDVIEW TASARIM
        // ----------------------------------------------------

        private void StyleDataGridView()
        {
            _dgvData.EnableHeadersVisualStyles = false;

            _dgvData.ColumnHeadersDefaultCellStyle.BackColor =
                Color.FromArgb(248, 249, 253);

            _dgvData.ColumnHeadersDefaultCellStyle.ForeColor =
                Color.FromArgb(65, 68, 90);

            _dgvData.ColumnHeadersDefaultCellStyle.Font =
                new Font(
                    "Segoe UI",
                    9,
                    FontStyle.Bold);

            _dgvData.ColumnHeadersHeight = 42;

            _dgvData.DefaultCellStyle.BackColor =
                Color.White;

            _dgvData.DefaultCellStyle.ForeColor =
                Color.FromArgb(65, 68, 90);

            _dgvData.DefaultCellStyle.SelectionBackColor =
                Color.FromArgb(225, 230, 255);

            _dgvData.DefaultCellStyle.SelectionForeColor =
                Color.FromArgb(45, 49, 70);

            _dgvData.DefaultCellStyle.Font =
                new Font("Segoe UI", 9);

            _dgvData.RowTemplate.Height = 34;

            _dgvData.GridColor =
                Color.FromArgb(235, 237, 243);

            _dgvData.CellBorderStyle =
                DataGridViewCellBorderStyle.SingleHorizontal;
        }

        private void ShowDynamicData()
        {
            _dgvData.DataSource = null;
            _dgvData.Columns.Clear();
            _dgvData.Rows.Clear();

            if (_kayitlar.Count == 0)
                return;

            // Excel'deki sütun başlıklarını tabloya ekle
            foreach (string header in _kayitlar[0].Values.Keys)
            {
                _dgvData.Columns.Add(header, header);
            }

            // Excel'deki satırları tabloya ekle
            foreach (ExcelDataRow row in _kayitlar)
            {
                object[] values = row.Values.Values
                    .Select(v => v is null ? string.Empty : (object)v)
                    .ToArray();
                _dgvData.Rows.Add(values);
            }

            _lblTotalValue.Text = _kayitlar.Count.ToString();
            _lblSuccessValue.Text = "-";
            _lblErrorValue.Text = "-";
        }
        private void ResetRowColors()
        {
            foreach (DataGridViewRow row in _dgvData.Rows)
            {
                row.DefaultCellStyle.BackColor = Color.White;
            }
        }

        // ----------------------------------------------------
        // BUTON
        // ----------------------------------------------------

        private Button CreateButton(
            string text,
            Color color)
        {
            Button button = new()
            {
                Text = text,
                Size = new Size(170, 42),

                BackColor = color,
                ForeColor = Color.White,

                FlatStyle = FlatStyle.Flat,

                Font = new Font(
                    "Segoe UI",
                    9,
                    FontStyle.Bold),

                Cursor = Cursors.Hand
            };

            button.FlatAppearance.BorderSize = 0;

            return button;
        }

        // ----------------------------------------------------
        // KPI KART
        // ----------------------------------------------------

        private Panel CreateCard(
            string title,
            string value,
            Color color,
            out Label valueLabel)
        {
            Panel panel = new()
            {
                Size = new Size(245, 100),
                BackColor = color
            };

            Label titleLabel = new()
            {
                Text = title,
                ForeColor =
                    Color.FromArgb(230, 230, 245),

                Font =
                    new Font("Segoe UI", 9),

                Location =
                    new Point(18, 15),

                AutoSize = true
            };

            valueLabel = new Label
            {
                Text = value,

                ForeColor = Color.White,

                Font = new Font(
                    "Segoe UI",
                    22,
                    FontStyle.Bold),

                Location =
                    new Point(16, 42),

                AutoSize = true
            };

            panel.Controls.Add(titleLabel);
            panel.Controls.Add(valueLabel);

            return panel;
        }

        // ----------------------------------------------------
        // PANEL
        // ----------------------------------------------------

        private Panel CreateWhitePanel()
        {
            return new Panel
            {
                BackColor = Color.White,
                BorderStyle = BorderStyle.None
            };
        }

        // ----------------------------------------------------
        // SIDEBAR BUTON
        // ----------------------------------------------------

        private void AddSideMenuButton(
            Panel sidebar,
            string text,
            int top)
        {
            Button button = new()
            {
                Text = text,

                Location =
                    new Point(12, top),

                Size =
                    new Size(50, 45),

                FlatStyle =
                    FlatStyle.Flat,

                BackColor =
                    Color.Transparent,

                ForeColor =
                    Color.White,

                Font =
                    new Font(
                        "Segoe UI",
                        14),

                Cursor =
                    Cursors.Hand
            };

            button.FlatAppearance.BorderSize = 0;

            sidebar.Controls.Add(button);
        }
    }

    // ========================================================
    // DONUT GRAFİĞİ
    // ========================================================

    public class DonutChart : Control
    {
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [DefaultValue(0d)]
        public double SuccessPercentage { get; set; } = 0;

        public DonutChart()
        {
            DoubleBuffered = true;
            BackColor = Color.White;
        }

        protected override void OnPaint(
            PaintEventArgs e)
        {
            base.OnPaint(e);

            e.Graphics.SmoothingMode =
                SmoothingMode.AntiAlias;

            int size =
                Math.Min(
                    Width,
                    Height) - 35;

            Rectangle rect = new(
                (Width - size) / 2,
                10,
                size,
                size);

            using Pen backgroundPen = new(
                Color.FromArgb(247, 89, 129),
                20);

            using Pen successPen = new(
                Color.FromArgb(72, 105, 246),
                20);

            backgroundPen.StartCap =
                LineCap.Round;

            backgroundPen.EndCap =
                LineCap.Round;

            successPen.StartCap =
                LineCap.Round;

            successPen.EndCap =
                LineCap.Round;

            e.Graphics.DrawArc(
                backgroundPen,
                rect,
                -90,
                360);

            float successAngle =
                (float)(360 *
                        SuccessPercentage / 100);

            if (successAngle > 0)
            {
                e.Graphics.DrawArc(
                    successPen,
                    rect,
                    -90,
                    successAngle);
            }

            string text =
                $"{SuccessPercentage:0}%";

            using Font percentageFont =
                new(
                    "Segoe UI",
                    18,
                    FontStyle.Bold);

            SizeF textSize =
                e.Graphics.MeasureString(
                    text,
                    percentageFont);

            e.Graphics.DrawString(
                text,
                percentageFont,
                Brushes.DimGray,
                (Width - textSize.Width) / 2,
                65);

            string description =
                "Başarı";

            using Font descFont =
                new(
                    "Segoe UI",
                    8);

            SizeF descSize =
                e.Graphics.MeasureString(
                    description,
                    descFont);

            e.Graphics.DrawString(
                description,
                descFont,
                Brushes.Gray,
                (Width - descSize.Width) / 2,
                100);
        }
       
        
    }
}