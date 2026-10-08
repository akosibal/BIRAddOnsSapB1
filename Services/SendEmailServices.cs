using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

public class SendEmailServices
{
    public SendEmailServices()
    {

    }

    public async Task<(bool result, string message)> SendReleasedDTC_ReportAsync(byte[] pdfBytes, string CardCode)
    {
        var recipients = "";/*  await _context.emailNotificationRecipientModels
                     .Where(u => u.active == true && u.trx_type == "ReleasedDTC_TRO")
                     .ToListAsync(); */


        //  var dtpDate = DateTime.Now.AddDays(-1);

        // 1️⃣ Generate PDF
        //  var pdfBytes = ""; /* await _excelExtraction.SendDailyReleasedDTC(dtpFrom, dtpTo); */

        // 2️⃣ Email body
        var subject = $"Released DTC Daily Reports";

        var body = $@"
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset='UTF-8'>
                    </head>
                    <body style='margin:0; padding:0; background:#f5f6f7; font-family:Arial, Helvetica, sans-serif;'>

                    <table width='100%' cellpadding='0' cellspacing='0' style='background:#f5f6f7;'>
                        <tr>
                            <td align='center' style='padding:20px;'>

                                <!-- MAIN CONTAINER -->
                                <table width='900' cellpadding='0' cellspacing='0' style='background:#ffffff;'>

                                    <!-- HEADER -->
                                    <tr>
                                        <td style='background:#EF9F2E; padding:8px 12px;'>
                                            <span style='color:#ffffff; font-size:26px; font-weight:bold;'>
                                                Released DTC Daily Reports
                                            </span>
                                        </td>
                                    </tr>

                                    <!-- CONTENT -->
                                    <tr>
                                        <td style='padding:16px 20px; color:#333333; font-size:18px;'>

                                            <p style='margin:0 0 14px;'>Good day,</p>

                                            <p style='margin:0 0 18px;'>
                                               Sending you the daily reports of the Released DTC.
                                            </p>

                                            <!-- INFO BOX -->
                                            <table width='100%' cellpadding='0' cellspacing='0'
                                                style='background:#f7f7f7; border-left:5px solid #198754; margin-bottom:18px;'>
                                                <tr>
                                                    <td style='padding:12px 16px;'>
                                                        <p style='margin:4px 0;'><strong>TRO Date:</strong> {DateTime.Now:MMMM dd, yyyy}</p>
                                                        <p style='margin:4px 0;'><strong>Generated Date:</strong> {DateTime.Now:MMMM dd, yyyy}</p>                                                       
                                                    </td>
                                                </tr>
                                            </table>

                                            <p style='margin-top:20px; margin-bottom:0px;'>
                                                The excel generated report is attached to this email.
                                            </p>                                           

                                            <p style='margin:0 0 30px;'>Thank you.</p>

                                            <hr style='border:none; border-top:1px solid #999;' />

                                            <p style='font-size:14px; color:#777777; font-style:italic; margin-top:12px;'>
                                                This is a system-generated email. Please do not reply.
                                            </p>

                                        </td>
                                    </tr>

                                </table>

                            </td>
                        </tr>
                    </table>

                    </body>
                    </html>";




        // 3️⃣ Send email

        var newEmailAttach = new List<EmailAttachment>()
        {
            new EmailAttachment()
            {
                  FileName = $"ReleasedDTC_{DateTime.Now.ToShortDateString()}.xlsx",
                   Content = pdfBytes
            }
        };

        var newEmail = new EmailAccnt()
        {
            //  To = "bacadimas@marsmandrysdale.com"//requestor.email
            // MltpleTo = recipients.Select(u => u.email).ToList()
        };

        try
        {
            await SendEmailAsync(
           newEmail,
           body,
            subject,
          newEmailAttach
       );
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }



        return (true, "Success");
    }


    private async Task SendEmailAsync(
      EmailAccnt email,
      string messageBody,
      string subject,
      List<EmailAttachment> attachments = null)
    {
        var message = new MimeMessage();

        /* =========================
           FROM
           ========================= */
        message.From.Add(MailboxAddress.Parse(email.CredUserId));

        /* =========================
           TO
           ========================= */
        if (!string.IsNullOrWhiteSpace(email.To))
        {
            message.To.Add(MailboxAddress.Parse(email.To));
        }
        else if (email.MltpleTo != null)
        {
            foreach (var to in email.MltpleTo.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                message.To.Add(MailboxAddress.Parse(to));
            }
        }

        /* =========================
           CC
           ========================= */
        if (!string.IsNullOrWhiteSpace(email.CC))
            message.Cc.Add(MailboxAddress.Parse(email.CC));

        if (email.MltpleCC != null)
        {
            foreach (var cc in email.MltpleCC.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                message.Cc.Add(MailboxAddress.Parse(cc));
            }
        }

        if (!string.IsNullOrWhiteSpace(email.GroupEmail))
            message.Cc.Add(MailboxAddress.Parse(email.GroupEmail));

        /* =========================
           BCC
           ========================= */
        if (!string.IsNullOrWhiteSpace(email.BCC))
            message.Bcc.Add(MailboxAddress.Parse(email.BCC));

        message.Subject = subject;

        /* =========================
           BODY + ATTACHMENTS
           ========================= */
        var builder = new BodyBuilder
        {
            HtmlBody = messageBody
        };

        if (attachments != null)
        {
            foreach (var att in attachments)
            {
                builder.Attachments.Add(
                    att.FileName,
                    att.Content,
                    ContentType.Parse(att.ContentType));
            }
        }

        message.Body = builder.ToMessageBody();

        /* =========================
           SEND
           ========================= */
        using var smtp = new SmtpClient();

        await smtp.ConnectAsync(
            email.Host,
            email.Port,
            email.EnableSsl ? SecureSocketOptions.StartTls : SecureSocketOptions.None);

        await smtp.AuthenticateAsync(email.CredUserId, email.Password);

        await smtp.SendAsync(message);
        await smtp.DisconnectAsync(true);
    }

}

public class EmailAccnt
{
    public string Host { get; set; } = "202.78.80.3";

    public int Port { get; set; } = 25;

    public bool EnableSsl { get; set; }

    public string CredUserId { get; set; } = "marsman@marsmandrysdale.com";

    public string Password { get; set; } = "marS1920man";

    public string To { get; set; }
    public List<string> MltpleTo { get; set; }

    public string CC { get; set; }
    public List<string> MltpleCC { get; set; }

    public string BCC { get; set; } = "bacadimas@marsmandrysdale.com";

    public string GroupEmail { get; set; }
}


public class EmailAttachment
{
    public string FileName { get; set; }
    public byte[] Content { get; set; }
    public string ContentType { get; set; } = "application/pdf";
}