using Raporlama.Services;

namespace Raporlama;

public partial class MailForm : Form
{
    private readonly string _attachmentPath;

    private TextBox txtSender = null!;
    private TextBox txtPassword = null!;
    private TextBox txtReceiver = null!;
    private TextBox txtSubject = null!;
    private TextBox txtBody = null!;

    public MailForm(string attachmentPath)
    {
        InitializeComponent();

        _attachmentPath = attachmentPath;

        BuildMailForm();
    }

    private void BuildMailForm()
    {
        Text = "Raporu E-posta ile Gönder";

        Size = new Size(520, 620);

        StartPosition = FormStartPosition.CenterParent;

        BackColor = Color.FromArgb(244, 246, 251);

        Font = new Font("Segoe UI", 9);

        FormBorderStyle = FormBorderStyle.FixedDialog;

        MaximizeBox = false;

        // Başlık
        Label title = new()
        {
            Text = "Rapor Gönderimi",
            Location = new Point(30, 20),
            AutoSize = true,

            Font = new Font(
                "Segoe UI",
                17,
                FontStyle.Bold),

            ForeColor = Color.FromArgb(38, 42, 66)
        };

        Controls.Add(title);


        // Gönderen
        txtSender = AddInput(
            "Gönderen Gmail Adresi",
            80);


        // Uygulama şifresi
        txtPassword = AddInput(
            "Gmail Uygulama Şifresi",
            145);

        txtPassword.UseSystemPasswordChar = true;


        // Alıcı
        txtReceiver = AddInput(
            "Alıcı E-posta Adresi",
            210);


        // Konu
        txtSubject = AddInput(
            "Konu",
            275);

        txtSubject.Text = "Üretim Raporu";


        // Mesaj
        Label bodyLabel = new()
        {
            Text = "Mesaj",
            Location = new Point(30, 340),
            AutoSize = true,
            ForeColor = Color.Gray
        };

        Controls.Add(bodyLabel);

        txtBody = new TextBox
        {
            Location = new Point(30, 365),

            Size = new Size(440, 90),

            Multiline = true,

            Text =
                "Merhaba,\r\n\r\n" +
                "Güncel üretim raporu ekte yer almaktadır.\r\n\r\n" +
                "İyi çalışmalar."
        };

        Controls.Add(txtBody);


        // Eklenen rapor
        Label attachmentLabel = new()
        {
            Text = $"Ek: {Path.GetFileName(_attachmentPath)}",

            Location = new Point(30, 470),

            AutoSize = true,

            ForeColor = Color.FromArgb(72, 105, 246)
        };

        Controls.Add(attachmentLabel);


        // Gönder butonu
        Button btnSend = new()
        {
            Text = "Mail Gönder",

            Location = new Point(300, 515),

            Size = new Size(170, 42),

            BackColor = Color.FromArgb(124, 58, 237),

            ForeColor = Color.White,

            FlatStyle = FlatStyle.Flat,

            Font = new Font(
                "Segoe UI",
                9,
                FontStyle.Bold),

            Cursor = Cursors.Hand
        };

        btnSend.FlatAppearance.BorderSize = 0;

        btnSend.Click += BtnSend_Click;

        Controls.Add(btnSend);
    }


    // ----------------------------------------------------
    // INPUT OLUŞTUR
    // ----------------------------------------------------

    private TextBox AddInput(
        string labelText,
        int top)
    {
        Label label = new()
        {
            Text = labelText,

            Location = new Point(30, top),

            AutoSize = true,

            ForeColor = Color.Gray
        };

        Controls.Add(label);

        TextBox textbox = new()
        {
            Location = new Point(30, top + 23),

            Size = new Size(440, 30)
        };

        Controls.Add(textbox);

        return textbox;
    }


    // ----------------------------------------------------
    // MAİL GÖNDER
    // ----------------------------------------------------

    private void BtnSend_Click(
        object? sender,
        EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtSender.Text) ||
            string.IsNullOrWhiteSpace(txtPassword.Text) ||
            string.IsNullOrWhiteSpace(txtReceiver.Text))
        {
            MessageBox.Show(
                "Gönderen, uygulama şifresi ve alıcı alanları zorunludur.",
                "Eksik Bilgi",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

            return;
        }

        try
        {
            Cursor = Cursors.WaitCursor;

            EmailService emailService = new();

            emailService.SendReport(
                txtSender.Text.Trim(),
                txtPassword.Text.Trim(),
                txtReceiver.Text.Trim(),
                txtSubject.Text.Trim(),
                txtBody.Text,
                _attachmentPath);

            MessageBox.Show(
                "Rapor başarıyla e-posta ile gönderildi.",
                "Başarılı",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message,
                "Mail Gönderme Hatası",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }
}