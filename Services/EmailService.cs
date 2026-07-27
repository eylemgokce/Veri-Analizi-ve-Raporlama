using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace Raporlama.Services;

public class EmailService
{
    public void SendReport(
        string senderEmail,
        string appPassword,
        string receiverEmail,
        string subject,
        string body,
        string attachmentPath)
    {
        if (!File.Exists(attachmentPath))
        {
            throw new FileNotFoundException(
                "Gönderilecek rapor dosyası bulunamadı.");
        }

        MimeMessage message = new();

        message.From.Add(
            MailboxAddress.Parse(senderEmail));

        message.To.Add(
            MailboxAddress.Parse(receiverEmail));

        message.Subject = subject;

        BodyBuilder bodyBuilder = new()
        {
            TextBody = body
        };

        bodyBuilder.Attachments.Add(attachmentPath);

        message.Body = bodyBuilder.ToMessageBody();

        using SmtpClient client = new();

        client.Connect(
            "smtp.gmail.com",
            587,
            SecureSocketOptions.StartTls);

        client.Authenticate(
            senderEmail,
            appPassword);

        client.Send(message);

        client.Disconnect(true);
    }
}