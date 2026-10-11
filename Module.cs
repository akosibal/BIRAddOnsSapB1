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
    string docType,
    string sent,
    string userCOde,
    DateTime InvoiceDate,
    string remarks)
    {
        /*    SAPbobsCOM.UserTable oUserTable = new SAPbobsCOM.UserTable(); */

        try
        {
            // 1. Get the UserTable object
            // Note: For UserTables.Item(), use the name WITHOUT the '@' prefix.
            /*  oUserTable = (SAPbobsCOM.UserTable)oCompany.GetBusinessObject(
                 SAPbobsCOM.BoObjectTypes.oUserTables); */

            SAPbobsCOM.UserTable oUserTable = oCompany.UserTables.Item("MDG_EIS_AUTOEMAIL");


            /*   oUserTable.UserFields.Fields.Item("DocEntry").Value = newDocEntry;
              oUserTable.UserFields.Fields.Item("LineId").Value = 0;
              oUserTable.UserFields.Fields.Item("VisOrder").Value = 0;
              oUserTable.UserFields.Fields.Item("Object").Value = "13"; // 13 = Sales Invoice
              oUserTable.UserFields.Fields.Item("LogInst").Value = 0; */

            // 4. Set your custom UDT fields (must start with "U_")
            oUserTable.UserFields.Fields.Item("U_InvNo").Value = docNum;
            oUserTable.UserFields.Fields.Item("U_Sent").Value = sent;
            oUserTable.UserFields.Fields.Item("U_CardCode").Value = cardCode;
            oUserTable.UserFields.Fields.Item("U_CardName").Value = cardName;
            oUserTable.UserFields.Fields.Item("U_DocNum").Value = newDocEntry;
            oUserTable.UserFields.Fields.Item("U_DocDate").Value = InvoiceDate.ToString("MM/dd/yyyy");
            oUserTable.UserFields.Fields.Item("U_DocType").Value = docType;
            oUserTable.UserFields.Fields.Item("U_Remarks").Value = remarks;
            oUserTable.UserFields.Fields.Item("U_UserCode").Value = userCOde;
            oUserTable.UserFields.Fields.Item("U_CreatedDate").Value = DateTime.Now.ToString("MM/dd/yyyy");

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
            /*  if (oUserTable != null)
             {
                 Marshal.ReleaseComObject(oUserTable);
                 oUserTable = null;
             } */
        }
    }


    public static void AddMenuItems(SAPbouiCOM.Application SBO_Application)
    {
        SAPbouiCOM.Menus oMenus = null;
        SAPbouiCOM.MenuItem oMenuItem = null;

        try
        {
            oMenus = SBO_Application.Menus;

            // ============================================
            // 1. CREATE THE PARENT FOLDER "Marsman Tools"
            // ============================================
            SAPbouiCOM.MenuCreationParams oFolder =
                (SAPbouiCOM.MenuCreationParams)SBO_Application.CreateObject(
                    SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

            oFolder.Type = SAPbouiCOM.BoMenuType.mt_POPUP;   // ← folder, not string
            oFolder.UniqueID = "MNU_MARSMAN_ROOT";
            oFolder.String = "Marsman Tools";
            oFolder.Position = 0;   // after all default folders

            // Attach at the TOP level of Main Menu
            oMenuItem = oMenus.Item("43520");   // Modules / Main Menu root
            oMenus = oMenuItem.SubMenus;
            oMenus.AddEx(oFolder);

            // ============================================
            // 2. ADD A CHILD ITEM INSIDE "Marsman Tools"
            // ============================================
            SAPbouiCOM.MenuCreationParams oChild =
                (SAPbouiCOM.MenuCreationParams)SBO_Application.CreateObject(
                    SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

            oChild.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            oChild.UniqueID = "MNU_AUTOEMAIL";
            oChild.String = "Auto Email Logs";
            oChild.Position = 0;

            // Get the folder we just created and add the child into it
            SAPbouiCOM.MenuItem oMarsmanFolder = oMenus.Item("MNU_MARSMAN_ROOT");
            SAPbouiCOM.Menus oFolderSub = oMarsmanFolder.SubMenus;
            oFolderSub.AddEx(oChild);

            Console.WriteLine("✅ Marsman Tools menu created with Auto Email Logs.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Menu error: {ex.Message}");
            SBO_Application.MessageBox($"Menu error: {ex.Message}");
        }
    }

}