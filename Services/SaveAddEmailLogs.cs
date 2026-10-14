using SAPbouiCOM;
using SAPbobsCOM;


public static class AddModules
{


    static Conn conn = new Conn();
    static SendEmailServices email = new SendEmailServices();


    public static void SaveDocument(SAPbouiCOM.Application SBO_Application,
                                    SAPbobsCOM.Company Company,
                                    ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo,
                                    ReportClass reportClass,
                                    string ObjectType,
                                    string DocumentType)
    {
        string docEntry = string.Empty;

        if (BusinessObjectInfo.EventType !=
        SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            return;

        // AFTER ADD
        if (BusinessObjectInfo.BeforeAction)
            return;

        Console.WriteLine("=================================");
        Console.WriteLine("Sales Invoice successfully added!");

        Console.WriteLine($"FormUID: {BusinessObjectInfo.FormUID}");
        Console.WriteLine($"ObjectKey: {BusinessObjectInfo.ObjectKey}");

        // ✅ CHECK 1: Global objects
        if (SBO_Application == null) { Console.WriteLine("NULL: SBO_Application"); return; }
        if (Company == null) { Console.WriteLine("NULL: Company"); return; }
        if (email == null) { Console.WriteLine("NULL: email"); return; }
        if (reportClass == null) { Console.WriteLine("NULL: reportClass"); return; }


        try
        {
            SAPbouiCOM.Form oForm =
                SBO_Application.Forms.Item(
                    BusinessObjectInfo.FormUID
                );

            if (oForm == null) { Console.WriteLine("NULL: oForm"); return; }

            Console.WriteLine($"Form Type: {oForm.TypeEx}");

            SAPbouiCOM.DBDataSource oDBDataSource = null;

            switch (ObjectType)
            {
                case "13":
                    oDBDataSource = oForm.DataSources.DBDataSources.Item("OINV");
                    break;
                case "14":
                    oDBDataSource = oForm.DataSources.DBDataSources.Item("ORIN");
                    break;
            }



      

            if (oDBDataSource == null) { Console.WriteLine("NULL: OINV DataSource"); return; }

            docEntry = oDBDataSource.GetValue("DocEntry", 0).Trim();
            string cardCode = oDBDataSource.GetValue("CardCode", 0).Trim();
            string cardName = oDBDataSource.GetValue("CardName", 0).Trim();
            string InvoiceNo = oDBDataSource.GetValue("DocNum", 0).Trim();
            string userSign = oDBDataSource.GetValue("UserSign", 0).Trim();

            string userCode = "";
            string userName = "";

            SAPbobsCOM.Recordset oRS = null;

            try
            {
                oRS = (SAPbobsCOM.Recordset)Company.GetBusinessObject(
                    SAPbobsCOM.BoObjectTypes.BoRecordset);



                oRS.DoQuery($@"
                            SELECT USER_CODE, U_NAME, E_Mail 
                            FROM OUSR 
                            WHERE USERID = {userSign}");

                if (!oRS.EoF)
                {
                    userCode = oRS.Fields.Item("USER_CODE").Value?.ToString() ?? "";
                    userName = oRS.Fields.Item("U_NAME").Value?.ToString() ?? "";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting user info: {ex.Message}");
            }
            finally
            {
                if (oRS != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            }


            string taxDateStr = oDBDataSource.GetValue("TaxDate", 0).Trim();

            DateTime taxDate;
            bool parsed = DateTime.TryParseExact(
                taxDateStr,
                "yyyyMMdd",                          // SAP B1 standard format
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out taxDate);

            if (!parsed)
            {
                // Fallback: try general parse if the format was different
                if (!DateTime.TryParse(taxDateStr, out taxDate))
                {
                    taxDate = DateTime.MinValue;  // or DateTime.Now if you prefer a default
                    Console.WriteLine($"Warning: Could not parse TaxDate '{taxDateStr}'");
                }
            }


            Console.WriteLine($"DocEntry: {docEntry}");

            if (string.IsNullOrEmpty(docEntry))
            {
                Console.WriteLine("DocEntry is empty.");

                SBO_Application.MessageBox(
                                "DocEntry is empty.",
                                1,
                                "OK",
                                "",
                                "");

                return;
            }
            else
            {

                var parameterValues = new Dictionary<string, object>
                {
                    { "dockey@",  docEntry },
                    { "objectid@", ObjectType },
                    { "usercode@", "manager" },
                };

                Console.WriteLine($"DocEntry: {docEntry}");
                var reportReulst = reportClass.GeneratePdf(Program.pathReportInvoice, parameterValues);

                var sendEmail = email.SendEInvoiceAsync(reportReulst, Company, cardCode, userSign, docEntry, DocumentType).GetAwaiter().GetResult();

                if (sendEmail.result)
                {
                    Module.SaveEmailLogs(docEntry,
                                    Company,
                                    SBO_Application,
                                    cardCode,        // Replace with actual CardCode
                                    cardName,// Replace with actual CardName
                                    InvoiceNo,    // Replace with actual DocNum
                                    ObjectType,          // DocType (13 = Invoice)
                                    "Y",
                                    userName,
                                    taxDate,
                                    "Email Sent Successfully"
                             );


                    SBO_Application.StatusBar.SetText(
                                $"Invoice {InvoiceNo}: Email sent successfully.",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Success
                                );
                }
                else
                {
                    // Email failed - do not save to UDT
                    Console.WriteLine($"Email failed: {sendEmail.message}");

                    Module.SaveEmailLogs(docEntry,
                                   Company,
                                   SBO_Application,
                                   cardCode,        // Replace with actual CardCode
                                   cardName,// Replace with actual CardName
                                   InvoiceNo,    // Replace with actual DocNum
                                   ObjectType,          // DocType (13 = Invoice)
                                   "N",
                                   userName,
                                   taxDate,
                                   "Failure to Send"
                            );

                    SBO_Application.MessageBox(
                        $"Email failed, data NOT saved to UDT:\n\n{sendEmail.message}",
                        1,
                        "OK",
                        "",
                        "");
                }
            }

            Console.WriteLine(
                $"Sales Invoice DocEntry = {docEntry}"
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"Error getting DocEntry: {ex}"
            );

            try
            {
                SBO_Application.MessageBox(
                    $"Error generating Crystal Report.\n\n" +
                    $"DocEntry: {docEntry}\n\n" +
                    $"Error:\n{ex.Message}",
                    1,
                    "OK",
                    "",
                    "");
            }
            catch
            {
                // Prevent MessageBox failure from crashing the add-on
            }
        }

    }
    public static void ShowAutoEmailForm(
          SAPbouiCOM.Application SBO_Application,
          SAPbobsCOM.Company oCompany)
    {
        try
        {
            // 1. If form already open, just focus it
            try
            {
                SAPbouiCOM.Form existing = SBO_Application.Forms.Item("BIR_FORM_1");
                existing.Visible = true;
                existing.Select();
                ApplyDefaultFiltersAndLoad(existing, oCompany, SBO_Application);
                return;
            }
            catch { /* not loaded — proceed */ }

            // 2. Create form via AddEx with FormCreationParams
            SAPbouiCOM.FormCreationParams cp =
                (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(
                    SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);

            cp.UniqueID = "BIR_FORM_1";
            cp.FormType = "BIR_FORM_1";
            //cp.ObjectType = "BIR_FORM_1";   // optional but helps
            cp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;

            SAPbouiCOM.Form oForm = SBO_Application.Forms.AddEx(cp);

            // 3. Configure form
            oForm.Title = "EIS Auto Email Logs";
            oForm.Left = 200;
            oForm.Top = 100;
            oForm.Width = 900;
            oForm.Height = 560;
            oForm.ClientWidth = 890;
            oForm.ClientHeight = 540;
            oForm.Visible = false;  // set visible after building

            // 4. Add items
            BuildFormItems(oForm);

            // 5. Show the form
            oForm.Visible = true;
            oForm.Freeze(false);

            // 6. Apply filters + load data


            // 4. Apply defaults + load data
            ApplyDefaultFiltersAndLoad(oForm, oCompany, SBO_Application);
        }
        catch (Exception ex)
        {
            SBO_Application.MessageBox(
                $"Error loading form: {ex.Message}",
                1, "Error", "", "");
        }
    }

    public static void ApplyDefaultFiltersAndLoad(
    SAPbouiCOM.Form oForm,
    SAPbobsCOM.Company oCompany,
    SAPbouiCOM.Application SBO_Application)
    {
        try
        {
            SAPbouiCOM.EditText edFrom = (SAPbouiCOM.EditText)oForm.Items.Item("edFrom").Specific;
            SAPbouiCOM.EditText edTo = (SAPbouiCOM.EditText)oForm.Items.Item("edTo").Specific;
            SAPbouiCOM.ComboBox cbSent = (SAPbouiCOM.ComboBox)oForm.Items.Item("cbSent").Specific;
            SAPbouiCOM.ComboBox cbDocType = (SAPbouiCOM.ComboBox)oForm.Items.Item("cbDocType").Specific;

            // Set date values in SAP's internal yyyyMMdd format
            edFrom.Value = DateTime.Now.AddDays(-30).ToString("yyyyMMdd");
            edTo.Value = DateTime.Now.ToString("yyyyMMdd");

            // Populate combo if empty
            if (cbSent.ValidValues.Count == 0)
            {
                cbSent.ValidValues.Add("N", "Failed to Send");
                cbSent.ValidValues.Add("Y", "Sent Successfully");
                cbSent.ValidValues.Add("", "All");
            }

            cbSent.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
            cbDocType.Select("13", SAPbouiCOM.BoSearchKey.psk_ByValue);

            LoadLogs(oForm, oCompany, SBO_Application);
        }
        catch (Exception ex)
        {
            SBO_Application.MessageBox(
                $"ApplyDefaultFilters error: {ex.Message}",
                1, "Error", "", "");
        }
    }

    public static void SendInvoiceEmail(
     SAPbouiCOM.Form oForm,
     SAPbobsCOM.Company oCompany,
     SAPbouiCOM.Application SBO_Application)
    {
        try
        {
            SAPbouiCOM.Matrix oMatrix =
                (SAPbouiCOM.Matrix)oForm.Items.Item("mtxLogs").Specific;

            var selectedInvoices = new List<string>();

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                SAPbouiCOM.CheckBox chk =
                    (SAPbouiCOM.CheckBox)oMatrix.Columns
                        .Item("colSelect").Cells.Item(i).Specific;

                if (!chk.Checked)
                    continue;

                SAPbouiCOM.EditText txt =
                    (SAPbouiCOM.EditText)oMatrix.Columns
                        .Item("colDN").Cells.Item(i).Specific;

                string docEntry = txt.Value.Trim();

                if (!string.IsNullOrEmpty(docEntry))
                    selectedInvoices.Add(docEntry);
            }

            if (selectedInvoices.Count == 0)
            {
                SBO_Application.StatusBar.SetText(
                    "Please select at least one invoice.",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                return;
            }

            int successCount = 0;
            int failedCount = 0;

            foreach (string docEntry in selectedInvoices)
            {
                try
                {
                    // Get invoice and customer details
                    SAPbobsCOM.Recordset rs =
                        (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                            SAPbobsCOM.BoObjectTypes.BoRecordset);

                    if (!int.TryParse(docEntry, out int invoiceDocEntry))
                        throw new Exception("Invalid invoice DocEntry: " + docEntry);

                    string sql = $@"
                    SELECT
                        T0.DocEntry,
                        T0.DocNum,
                        T0.CardCode,
                        T0.CardName,
                        T1.E_Mail AS EmailAddress,
                        T0.UserSign,
                        T0.TaxDate
                    FROM OINV T0
                    LEFT JOIN OCRD T1
                        ON T0.CardCode = T1.CardCode
                    WHERE T0.DocEntry = {invoiceDocEntry}";

                    rs.DoQuery(sql);

                    if (rs.EoF)
                        throw new Exception("Invoice not found: " + docEntry);

                    string cardCode = rs.Fields.Item("CardCode").Value.ToString();
                    string cardName = rs.Fields.Item("CardName").Value.ToString();
                    string emailAddress = rs.Fields.Item("EmailAddress").Value?.ToString() ?? "";
                    string docNum = rs.Fields.Item("DocNum").Value.ToString();
                    string userSign = rs.Fields.Item("UserSign").Value.ToString();

                    string taxDateStr = rs.Fields.Item("TaxDate").Value.ToString();

                    emailAddress = "bacadimas@marsmandrydale.com";

                    if (string.IsNullOrWhiteSpace(emailAddress))
                    {
                        throw new Exception(
                            $"Customer {cardCode} ({cardName}) has no email address.");
                    }

                    Console.WriteLine($"DocEntry: {docEntry}");
                    Console.WriteLine($"DocNum: {docNum}");
                    Console.WriteLine($"CardCode: {cardCode}");
                    Console.WriteLine($"CardName: {cardName}");
                    Console.WriteLine($"Email: {emailAddress}");

                    // Generate invoice PDF
                    var parameterValues = new Dictionary<string, object>
                {
                    { "dockey@", docEntry },
                    { "objectid@", "13" },
                    { "usercode@", "manager" }
                };

                    SAPbobsCOM.Recordset oRS = null;

                    string userCode = "";
                    string userName = "";

                    try
                    {
                        oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                            SAPbobsCOM.BoObjectTypes.BoRecordset);



                        oRS.DoQuery($@"
                            SELECT USER_CODE, U_NAME, E_Mail 
                            FROM OUSR 
                            WHERE USERID = {userSign}");

                        if (!oRS.EoF)
                        {
                            userCode = oRS.Fields.Item("USER_CODE").Value?.ToString() ?? "";
                            userName = oRS.Fields.Item("U_NAME").Value?.ToString() ?? "";
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error getting user info: {ex.Message}");
                    }
                    finally
                    {
                        if (oRS != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                    }

                    DateTime taxDate;
                    bool parsed = DateTime.TryParseExact(
                        taxDateStr,
                        "yyyyMMdd",                          // SAP B1 standard format
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out taxDate);

                    if (!parsed)
                    {
                        // Fallback: try general parse if the format was different
                        if (!DateTime.TryParse(taxDateStr, out taxDate))
                        {
                            taxDate = DateTime.MinValue;  // or DateTime.Now if you prefer a default
                            Console.WriteLine($"Warning: Could not parse TaxDate '{taxDateStr}'");
                        }
                    }



                    var reportResult = Program.reportClass.GeneratePdf(
                        Program.pathReportInvoice,
                        parameterValues);

                    // Send the PDF
                    var sendEmail = email.SendEInvoiceAsync(
                                                reportResult,
                                                oCompany,
                                                cardCode,
                                                userSign,
                                                docEntry,
                                                "A/R Invoice").GetAwaiter().GetResult();

                    if (sendEmail.result)
                    {
                        Module.UpdateEmailLogs(docEntry,
                                   oCompany,
                                   SBO_Application,
                                   cardCode,        // Replace with actual CardCode
                                   cardName,// Replace with actual CardName
                                   docNum,    // Replace with actual DocNum
                                   "13",          // DocType (13 = Invoice)
                                   "Y",
                                   userName,
                                   taxDate,
                                   "Email Sent Successfully"
                            );

                        SBO_Application.StatusBar.SetText(
                                    $"Invoice {docNum}: Email sent successfully.",
                                    SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Success
                                    );

                        successCount++;

                        AddModules.ShowAutoEmailForm(SBO_Application, oCompany);

                    }
                    else
                    {
                        Module.UpdateEmailLogs(docEntry,
                                   oCompany,
                                   SBO_Application,
                                   cardCode,        // Replace with actual CardCode
                                   cardName,// Replace with actual CardName
                                   docNum,    // Replace with actual DocNum
                                   "13",          // DocType (13 = Invoice)
                                   "N",
                                   userName,
                                   taxDate,
                                   sendEmail.message
                            );

                        failedCount++;

                        SBO_Application.StatusBar.SetText(
                            $"Invoice {docNum}: Failed to send email. {sendEmail.message}",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );

                    }
                }
                catch (Exception ex)
                {
                    failedCount++;

                    Console.WriteLine(
                        $"Invoice DocEntry {docEntry}: {ex.Message}");
                }
            }

            SBO_Application.MessageBox(
                $"Email process completed. Sent: {successCount}, Failed: {failedCount}.",
                1,                                         // 1 = OK button
                "Auto Email Result");                      // dialog title
        }
        catch (Exception ex)
        {
            SBO_Application.MessageBox(
                "Error sending invoice: " + ex.Message);
        }
    }

    public static void LoadLogs(
    SAPbouiCOM.Form oForm,
    SAPbobsCOM.Company oCompany,
    SAPbouiCOM.Application SBO_Application)
    {
        SAPbobsCOM.Recordset oRS = null;
        bool frozen = false;
        SAPbouiCOM.Matrix oMtx = null;

        try
        {
            Console.WriteLine("=== LoadLogs START ===");

            // ---- 1. Read filters from form ----
            SAPbouiCOM.EditText edFrom = (SAPbouiCOM.EditText)oForm.Items.Item("edFrom").Specific;
            SAPbouiCOM.EditText edTo = (SAPbouiCOM.EditText)oForm.Items.Item("edTo").Specific;
            SAPbouiCOM.ComboBox cbSent = (SAPbouiCOM.ComboBox)oForm.Items.Item("cbSent").Specific;
            SAPbouiCOM.ComboBox cbDocType = (SAPbouiCOM.ComboBox)oForm.Items.Item("cbDocType").Specific;

            string fromStr = edFrom.Value.Trim();
            string toStr = edTo.Value.Trim();
            string sentFilter = cbSent.Value;
            string docType = cbDocType.Value;

            DateTime dtFrom, dtTo;

            if (!DateTime.TryParseExact(fromStr, "yyyyMMdd",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out dtFrom))
            {
                dtFrom = DateTime.MinValue;
            }
            if (!DateTime.TryParseExact(toStr, "yyyyMMdd",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out dtTo))
            {
                dtTo = DateTime.MinValue;
            }

            string fromSql = dtFrom == DateTime.MinValue ? "" : dtFrom.ToString("yyyy-MM-dd");
            string toSql = dtTo == DateTime.MinValue ? "" : dtTo.ToString("yyyy-MM-dd");

            string fromYmd = dtFrom.ToShortDateString();
            string toYmd = dtTo.ToShortDateString();
            Console.WriteLine($"Normalized dates: from='{fromStr}' to='{toStr}'");

            Console.WriteLine($"Filters: from='{fromStr}' to='{toStr}' sent='{sentFilter}'");

            // ---- 2. Build WHERE clause ----
            string where = "1=1";
            if (!string.IsNullOrEmpty(sentFilter))
                where += $" AND U_Sent = '{sentFilter}'";
            if (!string.IsNullOrEmpty(fromStr))
                where += $" AND U_DocDate >= '{fromSql}'";
            if (!string.IsNullOrEmpty(toStr))
                where += $" AND U_DocDate <= '{toSql}'";

            if (!string.IsNullOrEmpty(docType))
                where += $" AND U_DocType = '{docType}'";

            Console.WriteLine($"WHERE: {where}");

            // ---- 3. Query the UDT ----
            oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset);

            var sql = $@"
            SELECT top 100 Code, Name, U_Sent, U_CardCode, U_CardName,
                   U_DocNum, U_DocDate, U_DocType, U_Remarks,
                   U_InvNo, U_UserCode, U_CreatedDate
            FROM [@MDG_EIS_AUTOEMAIL]
            WHERE {where}
            ORDER BY U_CreatedDate DESC";

            Console.Write("SQL {0}", sql);

            oRS.DoQuery(sql);

            Console.WriteLine($"Query returned: {oRS.RecordCount} rows");

            // ---- 4. Prepare matrix + UDS ----
            oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("mtxLogs").Specific;
            var uds = oForm.DataSources.UserDataSources;

            oForm.Freeze(true);
            frozen = true;

            // ---- 5. Clear existing rows ----
            Console.WriteLine($"Matrix before clear: RowCount = {oMtx.RowCount}");
            while (oMtx.RowCount > 0)
                oMtx.DeleteRow(1);
            oMtx.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;  // single row select

            Console.WriteLine($"Matrix after clear:  RowCount = {oMtx.RowCount}");

            // ---- 6. Populate rows ----
            int row = 0;
            while (!oRS.EoF)
            {
                row++;                    // ← row starts at 1

                // 1️⃣ Add the row FIRST
                oMtx.AddRow();            // ← no args — safest form
                Console.WriteLine($"Added row {row}, RowCount = {oMtx.RowCount}");

                // 2️⃣ Populate UDS with current row values
                uds.Item("UDS_Code").Value = GetStr(oRS, "Code");    // ← KEEP (hidden but populated)
                uds.Item("UDS_Select").Value = "N";
                // uds.Item("UDS_Name").Value = ...                    // ← REMOVE (Name column gone)
                uds.Item("UDS_Sent").Value = GetStr(oRS, "U_Sent");
                uds.Item("UDS_CC").Value = GetStr(oRS, "U_CardCode");
                uds.Item("UDS_CN").Value = GetStr(oRS, "U_CardName");
                uds.Item("UDS_DN").Value = GetStr(oRS, "U_DocNum");
                uds.Item("UDS_DD").Value = GetStr(oRS, "U_DocDate");
                uds.Item("UDS_DT").Value = GetStr(oRS, "U_DocType") == "13" ? "A/R INVOICE" : "CREDIT MEMO";
                uds.Item("UDS_IN").Value = GetStr(oRS, "U_InvNo");
                uds.Item("UDS_UC").Value = GetStr(oRS, "U_UserCode");
                uds.Item("UDS_CD").Value = GetStr(oRS, "U_CreatedDate");
                uds.Item("UDS_RM").Value = GetStr(oRS, "U_Remarks");

                // 3️⃣ Push UDS values into the visible row
                oMtx.SetLineData(row);

                oRS.MoveNext();
            }

            Console.WriteLine($"Rows added: {row}");
            Console.WriteLine($"Matrix RowCount after load: {oMtx.RowCount}");

            oForm.Freeze(false);
            frozen = false;

            Console.WriteLine("=== LoadLogs END ===");
        }
        catch (Exception ex)
        {
            if (frozen) oForm.Freeze(false);
            Console.WriteLine($"LoadLogs ERROR: {ex}");
            SBO_Application.MessageBox(
                $"LoadLogs error:\n{ex.Message}",
                1, "Error", "", "");
        }
        finally
        {
            if (oRS != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
        }
    }

    private static string GetStr(SAPbobsCOM.Recordset oRS, string fieldName)
    {
        object v = oRS.Fields.Item(fieldName).Value;
        return (v == null || v == DBNull.Value) ? "" : v.ToString();
    }


    private static void SetMatrixValue(
        SAPbouiCOM.Matrix oMtx,
        string colUid,
        int row,
        SAPbobsCOM.Field field)
    {
        object val = field.Value;

        string sval = (val == null || val == DBNull.Value)
            ? ""
            : Convert.ToString(
                val,
                System.Globalization.CultureInfo.InvariantCulture
              ) ?? "";

        if (!string.IsNullOrWhiteSpace(sval) &&
            (colUid == "colDD" || colUid == "colCD"))
        {
            if (DateTime.TryParse(
                sval,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out DateTime dt))
            {
                sval = dt.ToString("yyyy-MM-dd");
            }
        }

        SAPbouiCOM.EditText cell =
            (SAPbouiCOM.EditText)oMtx
                .Columns.Item(colUid)
                .Cells.Item(row)
                .Specific;

        cell.Value = sval;
    }

    private static void BuildFormItems(SAPbouiCOM.Form oForm)
    {
        SAPbouiCOM.Item oItem;
        SAPbouiCOM.StaticText oLbl;
        SAPbouiCOM.EditText oEdt;
        SAPbouiCOM.ComboBox oCbo;
        SAPbouiCOM.Button oBtn;
        SAPbouiCOM.Matrix oMtx;

        oForm.Freeze(true);

        // ============================================================
        // STEP 1: UserDataSources for filter fields (dates)
        // ============================================================
        var udsc = oForm.DataSources.UserDataSources;
        udsc.Add("UDS_FROM", SAPbouiCOM.BoDataType.dt_DATE);
        udsc.Add("UDS_TO", SAPbouiCOM.BoDataType.dt_DATE);

        // ============================================================
        // STEP 2: UserDataSources for matrix columns
        // Code is retained (hidden). Name is removed.
        // ============================================================
        udsc.Add("UDS_Code", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_Select", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
        udsc.Add("UDS_Sent", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_CC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_CN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_DN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_DD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_DT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_IN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_UC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_CD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
        udsc.Add("UDS_RM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);

        // ============================================================
        // STEP 3: Filter labels and fields
        // ============================================================

        // ---- lblFrom ----
        oItem = oForm.Items.Add("lblFrom", SAPbouiCOM.BoFormItemTypes.it_STATIC);
        oItem.Left = 10; oItem.Top = 12; oItem.Width = 40; oItem.Height = 16;
        ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "From:";

        // ---- edFrom ----
        oItem = oForm.Items.Add("edFrom", SAPbouiCOM.BoFormItemTypes.it_EDIT);
        oItem.Left = 55; oItem.Top = 10; oItem.Width = 90; oItem.Height = 20;
        oEdt = (SAPbouiCOM.EditText)oItem.Specific;
        oEdt.DataBind.SetBound(true, "", "UDS_FROM");

        // ---- lblTo ----
        oItem = oForm.Items.Add("lblTo", SAPbouiCOM.BoFormItemTypes.it_STATIC);
        oItem.Left = 160; oItem.Top = 12; oItem.Width = 25; oItem.Height = 16;
        ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "To:";

        // ---- edTo ----
        oItem = oForm.Items.Add("edTo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
        oItem.Left = 190; oItem.Top = 10; oItem.Width = 90; oItem.Height = 20;
        oEdt = (SAPbouiCOM.EditText)oItem.Specific;
        oEdt.DataBind.SetBound(true, "", "UDS_TO");

        // ---- lblSent ----
        oItem = oForm.Items.Add("lblSent", SAPbouiCOM.BoFormItemTypes.it_STATIC);
        oItem.Left = 300; oItem.Top = 12; oItem.Width = 40; oItem.Height = 16;
        ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Sent:";

        // ---- cbSent ----
        oItem = oForm.Items.Add("cbSent", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
        oItem.Left = 345; oItem.Top = 10; oItem.Width = 40; oItem.Height = 20;
        oCbo = (SAPbouiCOM.ComboBox)oItem.Specific;
        oCbo.ValidValues.Add("N", "Not Sent");
        oCbo.ValidValues.Add("Y", "Sent");
        oCbo.ValidValues.Add("", "All");

        // ---- lblDocType ----
        oItem = oForm.Items.Add("lblDocType", SAPbouiCOM.BoFormItemTypes.it_STATIC);
        oItem.Left = 420; oItem.Top = 12; oItem.Width = 60; oItem.Height = 16;
        ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Doc Type:";

        // ---- cbDocType ----
        oItem = oForm.Items.Add("cbDocType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
        oItem.Left = 480; oItem.Top = 10; oItem.Width = 50; oItem.Height = 20;
        oCbo = (SAPbouiCOM.ComboBox)oItem.Specific;
        oCbo.ValidValues.Add("13", "A/R Invoice");     // 13 = AR Invoice
        oCbo.ValidValues.Add("14", "Credit Memo");     // 14 = AR Credit Memo
        oCbo.ValidValues.Add("", "All");

        // ---- btnSearch ----
        oItem = oForm.Items.Add("btnSearch", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
        oItem.Left = 610; oItem.Top = 9; oItem.Width = 80; oItem.Height = 22;
        ((SAPbouiCOM.Button)oItem.Specific).Caption = "Search";

        oItem = oForm.Items.Add("btnSend", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
        oItem.Left = 800; oItem.Top = 9; oItem.Width = 80; oItem.Height = 22;
        ((SAPbouiCOM.Button)oItem.Specific).Caption = "Send";

        // ---- btnSendEmail (right edge) ----
        /*         oItem = oForm.Items.Add("btnSend", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 720; oItem.Top = 9; oItem.Width = 100; oItem.Height = 22;
                ((SAPbouiCOM.Button)oItem.Specific).Caption = "Send"; */

        // ---- btnClear ----
        /*   oItem = oForm.Items.Add("btnClear", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
          oItem.Left = 530; oItem.Top = 9; oItem.Width = 80; oItem.Height = 22;
          ((SAPbouiCOM.Button)oItem.Specific).Caption = "Clear"; */

        // ============================================================
        // STEP 4: Static header labels above the matrix
        // No "Code" header (hidden), no "Name" header (removed)
        // ============================================================
        /*   AddHeaderLabel(oForm, "hSelect", "Select", 10, 40);
          AddHeaderLabel(oForm, "hSent", "Sent", 50, 40);
          AddHeaderLabel(oForm, "hCC", "CardCode", 90, 70);
          AddHeaderLabel(oForm, "hCN", "CardName", 160, 130);
          AddHeaderLabel(oForm, "hDN", "DocNum", 290, 80);
          AddHeaderLabel(oForm, "hDD", "DocDate", 370, 80);
          AddHeaderLabel(oForm, "hDT", "DocType", 450, 100);
          AddHeaderLabel(oForm, "hIN", "InvNo", 550, 80);
          AddHeaderLabel(oForm, "hUC", "UserCode", 630, 70);
          AddHeaderLabel(oForm, "hCD", "Created", 700, 90);
          AddHeaderLabel(oForm, "hRM", "Remarks", 790, 150); */

        // ============================================================
        // STEP 5: Add matrix
        // ============================================================
        oItem = oForm.Items.Add("mtxLogs", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
        oItem.Left = 10; oItem.Top = 52;
        oItem.Width = 900; oItem.Height = 460;
        oMtx = (SAPbouiCOM.Matrix)oItem.Specific;

        // Hide built-in header row


        // ---- Code column (added first but hidden) ----
        AddHiddenMatrixColumn(oMtx, "colCode", 0, "UDS_Code");

        // ---- Visible columns ----
        AddCheckBoxMatrixColumn(oMtx, "colSelect", "UDS_Select", 40);
        AddMatrixColumn(oMtx, "colSent", 40, "UDS_Sent", "Sent");
        AddMatrixColumn(oMtx, "colCC", 70, "UDS_CC", "Customer COde");
        AddMatrixColumn(oMtx, "colCN", 130, "UDS_CN", "Customer Bane");
        AddMatrixColumn(oMtx, "colDN", 80, "UDS_DN", "Document Entry");
        AddMatrixColumn(oMtx, "colDD", 80, "UDS_DD", "Document Date");
        AddMatrixColumn(oMtx, "colDT", 100, "UDS_DT", "Document Type");
        AddMatrixColumn(oMtx, "colIN", 80, "UDS_IN", "Invoice No");
        AddMatrixColumn(oMtx, "colUC", 70, "UDS_UC", "Created By");
        AddMatrixColumn(oMtx, "colCD", 90, "UDS_CD", "Created Date");
        AddMatrixColumn(oMtx, "colRM", 150, "UDS_RM", "Remarks");

        oForm.Freeze(false);
    }

    private static void AddHiddenMatrixColumn(
    SAPbouiCOM.Matrix oMtx,
    string colUid,
    int width,
    string udsName)
    {
        SAPbouiCOM.IColumn col = oMtx.Columns.Add(colUid, SAPbouiCOM.BoFormItemTypes.it_EDIT);

        col.Width = width;      // can be 0 or any small value
        col.Editable = false;
        col.Visible = false;      // ← this hides the column

        col.DataBind.SetBound(true, "", udsName);
    }

    private static void AddCheckBoxMatrixColumn(
    SAPbouiCOM.Matrix oMtx,
    string colUid,
    string udsName,
    int width)
    {
        SAPbouiCOM.IColumn col = oMtx.Columns.Add(colUid, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);

        col.Width = width;
        col.Editable = true;     // ← must be editable to allow toggling
        col.Visible = true;

        // Bind to a UDS (text type, "Y"/"N")
        col.DataBind.SetBound(true, "", udsName);
    }


    private static void AddMatrixColumn(
      SAPbouiCOM.Matrix oMtx,
      string colUid,
      int width,
      string udsName,
      string title)
    {
        SAPbouiCOM.IColumn col = oMtx.Columns.Add(colUid, SAPbouiCOM.BoFormItemTypes.it_EDIT);

        col.Width = width;
        col.Editable = false;
        col.Visible = true;

        col.TitleObject.Caption = title;

        // ❌ Cannot set col.Title — it's read-only in this SDK
        // ✅ Header text is set via AddHeaderLabel() above the matrix

        // Bind to UDS
        col.DataBind.SetBound(true, "", udsName);
    }

    /*   private static void AddHeaderLabel(
          SAPbouiCOM.Form oForm,
          string uid,
          string caption,
          int left,
          int width)
      {
          SAPbouiCOM.Item oItem = oForm.Items.Add(uid, SAPbouiCOM.BoFormItemTypes.it_STATIC);
          oItem.Left = left;
          oItem.Top = 30;      // just above the matrix
          oItem.Width = width;
          oItem.Height = 18;

          SAPbouiCOM.StaticText oLbl = (SAPbouiCOM.StaticText)oItem.Specific;
          oLbl.Caption = caption;
      } */

    private static void AddHeaderLabel(
    SAPbouiCOM.Form oForm,
    string uid,
    string caption,
    int left,
    int width)
    {
        SAPbouiCOM.Item oItem = oForm.Items.Add(uid, SAPbouiCOM.BoFormItemTypes.it_STATIC);
        oItem.Left = left;
        oItem.Top = 28;    // moved up slightly so matrix border doesn't clip it
        oItem.Width = width;
        oItem.Height = 16;

        SAPbouiCOM.StaticText oLbl = (SAPbouiCOM.StaticText)oItem.Specific;
        oLbl.Caption = caption;
    }

}