
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using SAPbobsCOM;
using SAPbouiCOM;

public static class ExcelExporter
{
    /// <summary>
    /// Executes the SAP B1 query based on filters
    /// and exports the result to Excel.
    /// </summary>
    public static void ExportToExcel(
     SAPbobsCOM.Company oCompany,
     SAPbouiCOM.Application SBO_Application,
     DateTime dtFrom,
     DateTime dtTo,
     string sentFilter,
     string docType)
    {
        SAPbobsCOM.Recordset oRS = null;
        System.Windows.Forms.Cursor originalCursor = System.Windows.Forms.Cursor.Current;

        try
        {
            // 1. Prepare directory and RANDOM filename
            string directoryPath = "c:/temp";

            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                return; // Cancelled
            }

            Directory.CreateDirectory(directoryPath);

            // Generate a random filename (e.g., EIS_Export_a1b2c3d4-...xlsx)
            string randomFileName = $"EIS_Export_{Guid.NewGuid():N}.xlsx";
            string filePath = Path.Combine(directoryPath, randomFileName);

            // 2. Prepare filters
            string fromSql = dtFrom == DateTime.MinValue ? "" : dtFrom.ToString("yyyy-MM-dd");
            string toSql = dtTo == DateTime.MinValue ? "" : dtTo.ToString("yyyy-MM-dd");

            Console.WriteLine($"Filters: from='{fromSql}' to='{toSql}' sent='{sentFilter}' docType='{docType}'");

            // 3. Build WHERE clause
            string where = "1=1";

            if (!string.IsNullOrWhiteSpace(sentFilter))
            {
                string safeSent = sentFilter.Replace("'", "''");
                where += $" AND U_Sent = '{safeSent}'";
            }

            if (!string.IsNullOrWhiteSpace(fromSql))
            {
                where += $" AND U_DocDate >= '{fromSql}'";
            }

            if (!string.IsNullOrWhiteSpace(toSql))
            {
                where += $" AND U_DocDate < DATEADD(day, 1, '{toSql}')";
            }

            if (!string.IsNullOrWhiteSpace(docType))
            {
                string safeDocType = docType.Replace("'", "''");
                where += $" AND U_DocType = '{safeDocType}'";
            }

            Console.WriteLine($"WHERE: {where}");

            // 4. Execute SAP B1 query
            oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string sql = $@"
            SELECT TOP 100
                Code, Name, U_Sent, U_CardCode, U_CardName,
                U_DocNum, U_DocDate, U_DocType, U_Remarks,
                U_InvNo, U_UserCode, U_CreatedDate
            FROM [@MDG_EIS_AUTOEMAIL]
            WHERE {where}
            ORDER BY U_CreatedDate DESC";

            Console.WriteLine($"SQL: {sql}");
            oRS.DoQuery(sql);
            Console.WriteLine($"Query returned: {oRS.RecordCount} rows");

            // --- Show Wait Cursor & Status ---
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            SBO_Application.StatusBar.SetText("Generating Excel file...",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            // 5. Export to Excel using EPPlus
            ExcelPackage.License.SetNonCommercialPersonal("Dave");

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Email Logs");

                string[] headers = { "Code", "Name", "Sent", "Customer Code", "Customer Name",
                                 "Doc Entry", "Doc Date", "Doc Type", "Remarks",
                                 "Invoice No", "Created By", "Created Date" };

                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                }

                // Style headers
                using (ExcelRange headerRange = worksheet.Cells[1, 1, 1, headers.Length])
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                // 6. Write data rows
                int row = 2;
                int rowCounter = 0; // Counter for RemoveWindowsMessage

                while (!oRS.EoF)
                {
                    worksheet.Cells[row, 1].Value = GetFieldValue(oRS, "Code");
                    worksheet.Cells[row, 2].Value = GetFieldValue(oRS, "Name");
                    worksheet.Cells[row, 3].Value = GetFieldValue(oRS, "U_Sent");
                    worksheet.Cells[row, 4].Value = GetFieldValue(oRS, "U_CardCode");
                    worksheet.Cells[row, 5].Value = GetFieldValue(oRS, "U_CardName");
                    worksheet.Cells[row, 6].Value = GetFieldValue(oRS, "U_DocNum");

                    object docDate = oRS.Fields.Item("U_DocDate").Value;
                    if (docDate is DateTime documentDate && documentDate.Year >= 1900)
                    {
                        worksheet.Cells[row, 7].Value = documentDate;
                    }

                    worksheet.Cells[row, 8].Value = GetFieldValue(oRS, "U_DocType");
                    worksheet.Cells[row, 9].Value = GetFieldValue(oRS, "U_Remarks");
                    worksheet.Cells[row, 10].Value = GetFieldValue(oRS, "U_InvNo");
                    worksheet.Cells[row, 11].Value = GetFieldValue(oRS, "U_UserCode");

                    object createdDate = oRS.Fields.Item("U_CreatedDate").Value;
                    if (createdDate is DateTime created && created.Year >= 1900)
                    {
                        worksheet.Cells[row, 12].Value = created;
                    }

                    oRS.MoveNext();
                    row++;
                    rowCounter++;

                    // Allow SAP to process messages every 50 rows to prevent RPC_E_SYS_CALL_FAILED
                    if (rowCounter % 50 == 0)
                    {
                        SBO_Application.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, true);
                    }
                }

                // 7. Final formatting
                worksheet.View.FreezePanes(2, 1);

                worksheet.Cells[1, 1, Math.Max(row - 1, 1), headers.Length].AutoFilter = true;
                worksheet.Cells[1, 1, Math.Max(row - 1, 1), headers.Length].AutoFitColumns();

                worksheet.Column(7).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(12).Style.Numberformat.Format = "mm/dd/yyyy";

                // Final message pump before heavy save
                SBO_Application.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, true);

                // 8. Save the Excel file
                FileInfo fileInfo = new FileInfo(filePath);
                package.SaveAs(fileInfo);
            }

            SBO_Application.StatusBar.SetText(
                                            $"Excel export successful: {filePath}",
                                            SAPbouiCOM.BoMessageTime.bmt_Short,
                                            SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            // --- ADD THIS LINE ---
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Excel export error: {ex}");

            try
            {
                SBO_Application.StatusBar.SetText(
                    $"Excel export failed: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            catch { /* Status bar reporting may also fail */ }
        }
        finally
        {
            // Restore Cursor
            System.Windows.Forms.Cursor.Current = originalCursor;

            // 10. Release SAP COM object
            if (oRS != null)
            {
                try
                {
                    if (Marshal.IsComObject(oRS))
                    {
                        Marshal.ReleaseComObject(oRS);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Recordset cleanup error: {ex.Message}");
                }
                oRS = null;
            }
        }


    }

    /// <summary>
    /// Safely retrieves a field value as text.
    /// </summary>
    private static string GetFieldValue(
        SAPbobsCOM.Recordset recordset,
        string fieldName)
    {
        object value = recordset.Fields.Item(fieldName).Value;

        if (value == null ||
            value == DBNull.Value)
        {
            return string.Empty;
        }

        return value.ToString();
    }


    private static string GetExcelSavePath()
    {
        string filePath = null;

        Thread t = new Thread(() =>
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Title = "Export Email Logs to Excel";
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.DefaultExt = "xlsx";
                sfd.AddExtension = true;
                sfd.OverwritePrompt = true;
                sfd.FileName =
                    $"EIS_Email_Logs_{DateTime.Now:yyyyMMdd}.xlsx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    filePath = sfd.FileName;
                }
            }
        });

        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();

        return filePath;
    }

}
