namespace Raporlama;


partial class frmMain
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        btnSelectExcel = new Button();
        btnAnalyze = new Button();
        lbl_total = new Label();
        lbl_success = new Label();
        lbl_error = new Label();
        dgv_data = new DataGridView();
        openFileDialog1 = new OpenFileDialog();
        statusStrip1 = new StatusStrip();
        lbl_status = new ToolStripStatusLabel();
        statusStrip2 = new StatusStrip();
        lbl_title = new Label();
        ((System.ComponentModel.ISupportInitialize)dgv_data).BeginInit();
        statusStrip1.SuspendLayout();
        SuspendLayout();
        // 
        // btnSelectExcel
        // 
        btnSelectExcel.Cursor = Cursors.Hand;
        btnSelectExcel.FlatAppearance.BorderSize = 0;
        btnSelectExcel.FlatStyle = FlatStyle.Flat;
        btnSelectExcel.Location = new Point(82, 126);
        btnSelectExcel.Name = "btnSelectExcel";
        btnSelectExcel.Size = new Size(190, 42);
        btnSelectExcel.TabIndex = 1;
        btnSelectExcel.Text = "Excel dosyası seç";
        btnSelectExcel.UseVisualStyleBackColor = true;
        btnSelectExcel.Click += btnSelectExcel_Click;
        // 
        // btnAnalyze
        // 
        btnAnalyze.Cursor = Cursors.Hand;
        btnAnalyze.FlatAppearance.BorderSize = 0;
        btnAnalyze.FlatStyle = FlatStyle.Flat;
        btnAnalyze.Location = new Point(388, 126);
        btnAnalyze.Name = "btnAnalyze";
        btnAnalyze.Size = new Size(190, 42);
        btnAnalyze.TabIndex = 2;
        btnAnalyze.Text = "Analiz Et";
        btnAnalyze.UseVisualStyleBackColor = true;
        btnAnalyze.Click += btnAnalyze_Click;
        // 
        // lbl_total
        // 
        lbl_total.AutoSize = true;
        lbl_total.Location = new Point(82, 218);
        lbl_total.Name = "lbl_total";
        lbl_total.Size = new Size(130, 23);
        lbl_total.TabIndex = 3;
        lbl_total.Text = "Toplam Kayıt : 0";
        // 
        // lbl_success
        // 
        lbl_success.AutoSize = true;
        lbl_success.Location = new Point(82, 282);
        lbl_success.Name = "lbl_success";
        lbl_success.Size = new Size(86, 23);
        lbl_success.TabIndex = 4;
        lbl_success.Text = "Başarılı : 0";
        // 
        // lbl_error
        // 
        lbl_error.AutoSize = true;
        lbl_error.Location = new Point(82, 347);
        lbl_error.Name = "lbl_error";
        lbl_error.Size = new Size(77, 23);
        lbl_error.TabIndex = 5;
        lbl_error.Text = "Hatalı : 0";
        // 
        // dgv_data
        // 
        dgv_data.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgv_data.Location = new Point(80, 401);
        dgv_data.Name = "dgv_data";
        dgv_data.RowHeadersWidth = 51;
        dgv_data.Size = new Size(930, 216);
        dgv_data.TabIndex = 6;
        // 
        // openFileDialog1
        // 
        openFileDialog1.FileName = "openFileDialog1";
        // 
        // statusStrip1
        // 
        statusStrip1.ImageScalingSize = new Size(20, 20);
        statusStrip1.Items.AddRange(new ToolStripItem[] { lbl_status });
        statusStrip1.Location = new Point(0, 725);
        statusStrip1.Name = "statusStrip1";
        statusStrip1.Padding = new Padding(1, 0, 16, 0);
        statusStrip1.Size = new Size(1217, 26);
        statusStrip1.TabIndex = 7;
        statusStrip1.Text = "statusStrip1";
        statusStrip1.ItemClicked += statusStrip1_ItemClicked;
        // 
        // lbl_status
        // 
        lbl_status.Name = "lbl_status";
        lbl_status.Size = new Size(44, 20);
        lbl_status.Text = "Hazır";
        // 
        // statusStrip2
        // 
        statusStrip2.ImageScalingSize = new Size(20, 20);
        statusStrip2.Location = new Point(0, 703);
        statusStrip2.Name = "statusStrip2";
        statusStrip2.Padding = new Padding(1, 0, 16, 0);
        statusStrip2.Size = new Size(1217, 22);
        statusStrip2.TabIndex = 8;
        statusStrip2.Text = "statusStrip2";
        // 
        // lbl_title
        // 
        lbl_title.AutoSize = true;
        lbl_title.Location = new Point(166, 40);
        lbl_title.Name = "lbl_title";
        lbl_title.Size = new Size(354, 23);
        lbl_title.TabIndex = 0;
        lbl_title.Text = "Üretim Raporlama ve Veri Doğrulama Sistemi";
        // 
        // frmMain
        // 
        AutoScaleDimensions = new SizeF(9F, 23F);
        AutoScaleMode = AutoScaleMode.Font;
        BackColor = Color.White;
        ClientSize = new Size(1217, 751);
        Controls.Add(lbl_title);
        Controls.Add(statusStrip2);
        Controls.Add(statusStrip1);
        Controls.Add(dgv_data);
        Controls.Add(lbl_error);
        Controls.Add(lbl_success);
        Controls.Add(lbl_total);
        Controls.Add(btnAnalyze);
        Controls.Add(btnSelectExcel);
        Font = new Font("Segoe UI", 10.2F, FontStyle.Regular, GraphicsUnit.Point, 162);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        Name = "frmMain";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Üretim Raporlama ve Veri Doğrulama Sistemi";
        Load += frmMain_Load;
        ((System.ComponentModel.ISupportInitialize)dgv_data).EndInit();
        statusStrip1.ResumeLayout(false);
        statusStrip1.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
    private Button btnSelectExcel;
    private Button btnAnalyze;
    private Label lbl_total;
    private Label lbl_success;
    private Label lbl_error;
    private DataGridView dgv_data;
    private OpenFileDialog openFileDialog1;
    private StatusStrip statusStrip1;
    private ToolStripStatusLabel lbl_status;
    private StatusStrip statusStrip2;
    private Label lbl_title;
}
