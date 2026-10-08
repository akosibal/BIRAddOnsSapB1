using SAPbobsCOM;
using System;
using System.Runtime.InteropServices;

public static class Module
{


    public static bool SaveEmailLogs(
        string newDocEntry,
    SAPbobsCOM.Company oCompany,
    SAPbouiCOM.Application sboApplication,
    string cardCode,
    string cardName,
    string docNum,
    string docType)
    {
        SAPbobsCOM.UserTable oUserTable = null;

        try
        {
            // 1. Get the UserTable object
            // Note: For UserTables.Item(), use the name WITHOUT the '@' prefix.
            oUserTable = (SAPbobsCOM.UserTable)oCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.oUserTables);

            oUserTable = oCompany.UserTables.Item("EIS_AUTOEMAIL");

          
            oUserTable.UserFields.Fields.Item("DocEntry").Value = newDocEntry;
            oUserTable.UserFields.Fields.Item("LineId").Value = 0;
            oUserTable.UserFields.Fields.Item("VisOrder").Value = 0;
            oUserTable.UserFields.Fields.Item("Object").Value = "13"; // 13 = Sales Invoice
            oUserTable.UserFields.Fields.Item("LogInst").Value = 0;

            // 4. Set your custom UDT fields (must start with "U_")
            oUserTable.UserFields.Fields.Item("U_Sent").Value = "Y";
            oUserTable.UserFields.Fields.Item("U_CardCode").Value = cardCode;
            oUserTable.UserFields.Fields.Item("U_CardName").Value = cardName;
            oUserTable.UserFields.Fields.Item("U_DocNum").Value = docNum;
            oUserTable.UserFields.Fields.Item("U_DocDate").Value = DateTime.Now;
            oUserTable.UserFields.Fields.Item("U_DocType").Value = docType;
            oUserTable.UserFields.Fields.Item("U_Remarks").Value = "Email sent successfully";

            // 5. Add the record to the database
            int retCode = oUserTable.Add();

            if (retCode != 0)
            {
                string errMsg;
                oCompany.GetLastError(out retCode, out errMsg);

                Console.WriteLine($"Error saving to UDT: {errMsg}");
                sboApplication.MessageBox(
                    $"Email sent, but failed to save to UDT:\n\n{errMsg}",
                    1, "OK", "", "");
                return false;
            }

            Console.WriteLine($"Successfully saved log to UDT with DocEntry {newDocEntry}.");
            return true;
        }
        catch (Exception exUdt)
        {
            Console.WriteLine($"Exception saving to UDT: {exUdt.Message}");
            sboApplication.MessageBox(
                $"Email sent, but UDT save failed:\n\n{exUdt.Message}",
                1, "OK", "", "");
            return false;
        }
        finally
        {
            if (oUserTable != null)
            {
                Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
            }
        }
    }
}