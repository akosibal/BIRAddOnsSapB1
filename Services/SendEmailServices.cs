using System.Globalization;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

public class SendEmailServices
{
    public SendEmailServices()
    {

    }

    public async Task<(bool result, string message)> SendEInvoiceAsync(
     byte[] pdfBytes,
     SAPbobsCOM.Company oCompany,
     string CardCode,
     string userSign,
     string docEntry,
     string ObjectType,
     string docType = "Electronic Sales Invoice")
    {
        SAPbobsCOM.BusinessPartners oBP = null;
        SAPbobsCOM.Recordset oRS = null;
        var recipients = new List<string>();
        var userEmail = new List<string>();

        // Placeholders for template
        string customerName = "";
        string docNum = "";
        string docDate = "";
        string dueDate = "";
        string currency = "";
        string totalAmount = "";
        string paymentTerms = "";
        string userName = "";
        string companyName = "MEPI";           // change or pull from OADM
        string companyContact = "123-4567";       // change or pull from config

        try
        {

            oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRS.DoQuery(@"
                    SELECT 
                        CompnyName,
                        Phone1,
                        Phone2,
                        Fax,
                        E_Mail,
                        CompnyAddr,
                        State,
                        Country
                    FROM OADM");

            if (!oRS.EoF)
            {
                companyName = oRS.Fields.Item("CompnyName").Value?.ToString() ?? "";
                companyContact = oRS.Fields.Item("Phone1").Value?.ToString() ?? "";

                // Optional fallback to Phone2 if Phone1 is empty
                if (string.IsNullOrWhiteSpace(companyContact))
                {
                    companyContact = oRS.Fields.Item("Phone2").Value?.ToString() ?? "";
                }
            }

            // ============================
            // 1. Customer (email + name) from OCRD
            // ============================
            oBP = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.oBusinessPartners);

            if (oBP.GetByKey(CardCode))
            {
                customerName = oBP.CardName ?? "";

                if (!string.IsNullOrWhiteSpace(oBP.EmailAddress))
                {
                    recipients.AddRange(
                        oBP.EmailAddress
                            .Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(e => e.Trim())
                            .Where(e => !string.IsNullOrWhiteSpace(e))
                    );
                }
            }

            recipients.Clear();

            if (!recipients.Any())
            {
                recipients.Add("bacadimas@marsmandrysdale.com");
            }

            // ============================
            // 2. Invoice details from OINV
            // ============================
            oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset);

            switch (ObjectType)
            {
                case "13":   // A/R Invoice
                    oRS.DoQuery($@"
                                SELECT 
                                    T0.DocNum,
                                    T0.DocDate,
                                    T0.DocDueDate,
                                    T0.DocTotal,
                                    T0.DocCur,
                                    T0.PaymentRef,
                                    T1.PymntGroup AS PaymentTerms,
                                    T0.UserSign,
                                    T0.NumAtCard
                                FROM OINV T0
                                LEFT JOIN OCTG T1 ON T0.GroupNum = T1.GroupNum
                                WHERE T0.DocEntry = {docEntry}");
                    break;

                case "14":   // A/R Credit Memo
                    oRS.DoQuery($@"
                                SELECT 
                                    T0.DocNum,
                                    T0.DocDate,
                                    T0.DocDueDate,
                                    T0.DocTotal,
                                    T0.DocCur,
                                    T0.PaymentRef,
                                    T1.PymntGroup AS PaymentTerms,
                                    T0.UserSign,
                                    T0.NumAtCard
                                FROM ORIN T0
                                LEFT JOIN OCTG T1 ON T0.GroupNum = T1.GroupNum
                                WHERE T0.DocEntry = {docEntry}");
                    break;
            }



            if (!oRS.EoF)
            {

                docDate = Convert.ToDateTime(oRS.Fields.Item("DocDate").Value).ToString("MMMM dd, yyyy");
                dueDate = Convert.ToDateTime(oRS.Fields.Item("DocDueDate").Value).ToString("MMMM dd, yyyy");
                currency = oRS.Fields.Item("DocCur").Value?.ToString() ?? "";
                totalAmount = Convert.ToDecimal(oRS.Fields.Item("DocTotal").Value).ToString("N2");
                paymentTerms = oRS.Fields.Item("PaymentTerms").Value?.ToString() ?? "";
                docNum = oRS.Fields.Item("DocNum").Value?.ToString() ?? "";
            }

            // ============================
            // 3. SAP user (name + email) from OUSR
            // ============================
            oRS.DoQuery($"SELECT U_NAME, E_Mail FROM OUSR WHERE USERID = {userSign}");

            if (!oRS.EoF)
            {
                userName = oRS.Fields.Item("U_NAME").Value?.ToString() ?? "";

                var user1Email = oRS.Fields.Item("E_Mail").Value?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(user1Email))
                {
                    userEmail.AddRange(
                        user1Email
                            .Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(e => e.Trim())
                            .Where(e => !string.IsNullOrWhiteSpace(e))
                    );
                }
            }

            Console.WriteLine($"Recipients: {string.Join(", ", recipients)}");
        }
        catch (Exception ex)
        {
            return (false, $"Error building email data: {ex.Message}");
        }
        finally
        {
            if (oBP != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP);
            if (oRS != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
        }

        if (recipients.Count == 0)
            return (false, "No email recipients found.");

        // ============================
        // 4. Subject (new format)
        // ============================
        var subject = $"NO Reply {companyName}_{docType}_{docNum}";
        var documentName = ObjectType switch
                        {
                            "13" => "Invoice No",
                            "14" => "Credit Memo No",
                            _ => "Document No"
                        };


        // ============================
        // 5. Body (new template)
        // ============================
        var body = $@"
        <!DOCTYPE html>
        <html>
        <head><meta charset='UTF-8'></head>
        <body style='margin:0; padding:0; background:#f5f6f7; font-family:Arial, Helvetica, sans-serif;'>

        <table width='100%' cellpadding='0' cellspacing='0' style='background:#f5f6f7;'>
            <tr>
                <td align='center' style='padding:20px;'>
                    <table width='900' cellpadding='0' cellspacing='0' style='background:#ffffff;'>
                
                        <!-- CONTENT -->
                        <tr>
                            <td style='padding:16px 20px; color:#333333; font-size:18px;'>

                                <p style='margin:0 0 14px;'>Dear {CultureInfo.CurrentCulture.TextInfo.ToTitleCase(customerName.ToLower())},</p>

                                <p style='margin:0 0 18px;'>
                                    Please see attached {docType} No. <strong>{docNum}</strong> 
                                    for the products or services dated <strong>{docDate}</strong> 
                                    with the following details:
                                </p>

                                <!-- INFO BOX -->
                                <table width='100%' cellpadding='0' cellspacing='0'
                                    style='background:#f7f7f7; border-left:5px solid #198754; margin-bottom:18px;'>
                                    <tr>
                                        <td style='padding:12px 16px;'>
                                            <p style='margin:4px 0;'><strong>{documentName}:</strong> {docNum}</p>
                                            <p style='margin:4px 0;'><strong>Total Amount:</strong> {currency} {totalAmount}</p>
                                            <p style='margin:4px 0;'><strong>Payment Due Date:</strong> {dueDate}</p>
                                            <p style='margin:4px 0;'><strong>Payment Terms:</strong> {paymentTerms}</p>
                                        </td>
                                    </tr>
                                </table>

                                <p style='margin:10px 0 0 18px;'>
                                    You can submit your payment via our standard payment channels.
                                </p>

                                <p style='margin:0 0 18px;'>
                                    Please do not reply to this email. Should you have any concerns, 
                                    kindly contact <strong>{companyContact}</strong> and look for 
                                    <strong>{userName}</strong>.
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

        // ============================
        // 6. Attachment + Email object
        // ============================
        var newEmailAttach = new List<EmailAttachment>()
    {
        new EmailAttachment()
        {
            FileName = $"{docType.Replace(" ", "_")}_{docNum}.pdf",
            Content = pdfBytes
        }
    };

        var newEmail = new EmailAccnt()
        {

            MltpleTo = recipients,
            BCC = userEmail.FirstOrDefault()
        };

        // ============================
        // 7. Send
        // ============================
        try
        {
            await SendEmailAsync(newEmail, body, subject, newEmailAttach);
            return (true, "Success");
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
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