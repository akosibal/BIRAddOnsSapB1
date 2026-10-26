using SAPbouiCOM;
using SAPbobsCOM;
using Microsoft.Extensions.Configuration;


class Program
{

    public static IConfiguration Configuration { get; private set; }

    private static SAPbouiCOM.Application SBO_Application;
    private static SAPbobsCOM.Company Company;

    static Conn conn = new Conn();
    static SendEmailServices email = new SendEmailServices();

    public static ReportClass reportClass = new ReportClass("192.168.5.22", "TEST_MDTI_20260732", "sa", "1q2w#E$R");
    public static string pathReportTemplates { get; set; } = "";


    static void Main(string[] args)
    {
          Console.WriteLine("SAP Business One Add-on");

        if (args.Length == 0)
        {
              Console.WriteLine(
                  "SAP Business One connection string was not provided."
              );
   
            return;
        }

        try
        {
            var builder = new ConfigurationBuilder()
              // Use BaseDirectory to ensure it looks in the bin folder, NOT the SAP folder
              .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
              .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            Configuration = builder.Build();

            reportClass = new ReportClass(Configuration["SapConnection:Server"],Configuration["SapConnection:CompanyDB"],Configuration["SapConnection:UserName"],Configuration["SapConnection:Password"]);
            var conn = new Conn();

            conn.Connect(args[0]);

            // Reuse the existing connections
            SBO_Application = conn.SBO_Application;
            Company = conn.company;

            // Console.WriteLine("Add-on started.");

            SBO_Application.MenuEvent += SBO_Application_MenuEvent;
            SBO_Application.FormDataEvent +=
                SBO_Application_FormDataEvent;

            SBO_Application.AppEvent += SBO_Application_AppEvent;
            SBO_Application.ItemEvent += SBO_Application_ItemEvent;

            Module.AddMenuItems(SBO_Application);


            //  Console.WriteLine("FormDataEvent registered.");
              Console.WriteLine("Waiting for SAP Business One events...");

            /*       Module.SaveEmailLogs("1",
                                       Company,
                                       SBO_Application,
                                       "C0001",        // Replace with actual CardCode
                                       "Customer Name",// Replace with actual CardName
                                       "INV-12345",    // Replace with actual DocNum
                                       "INVOICE"            // DocType (13 = Invoice)
                                ); */


            System.Threading.Thread.Sleep(
               System.Threading.Timeout.Infinite
           );
        }
        catch (Exception ex)
        {
              Console.WriteLine(ex);
        }
    }

    static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
    {
        switch (EventType)
        {
            case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
            case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
            case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                RemoveMenuItems();
                System.Windows.Forms.Application.Exit();
                break;
        }
    }

    private static void RemoveMenuItems()
    {
        try
        {
            SBO_Application.Menus.RemoveEx("MNU_MARSMAN_ROOT");
        }
        catch { /* ignore */ }
    }

    static void SBO_Application_ItemEvent(
     string FormUID,
     ref SAPbouiCOM.ItemEvent pVal,
     out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            // Only handle our form
            if (FormUID != "BIR_FORM_1") return;

            // Only handle button clicks (after action)
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) return;
            if (pVal.BeforeAction) return;

            var oForm = SBO_Application.Forms.Item("BIR_FORM_1");

            switch (pVal.ItemUID)
            {
                case "btnSearch":
                    try
                    {
                        System.Windows.Forms.Cursor.Current =
                             System.Windows.Forms.Cursors.WaitCursor;

                        Console.WriteLine("Search clicked");
                        AddModules.LoadLogs(oForm, Company, SBO_Application);
                    }
                    finally
                    {
                        System.Windows.Forms.Cursor.Current =
                            System.Windows.Forms.Cursors.Default;
                    }

                    break;

                case "btnClear":
                    Console.WriteLine("Clear clicked");
                    AddModules.ApplyDefaultFiltersAndLoad(oForm, Company, SBO_Application);
                    break;

                case "btnSend":
                    int response = SBO_Application.MessageBox(
                                    "Are you sure you want to send the selected invoice emails?",
                                    2,
                                    "Yes",
                                    "No"
                                );

                    if (response == 1)
                    {
                        try
                        {
                            System.Windows.Forms.Cursor.Current =
                                System.Windows.Forms.Cursors.WaitCursor;

                            SBO_Application.StatusBar.SetText(
                                "Sending invoice emails, please wait...",
                                SAPbouiCOM.BoMessageTime.bmt_Long,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                            );

                            AddModules.SendInvoiceEmail(
                                oForm,
                                Company,
                                SBO_Application
                            );
                        }
                        finally
                        {
                            System.Windows.Forms.Cursor.Current =
                                System.Windows.Forms.Cursors.Default;
                        }
                    }
                    break;

                case "btnExport":
                    try
                    {
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                        SBO_Application.StatusBar.SetText("Exporting to Excel, please wait...",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        // 1. Get the string filters from the SAP Form
                        var filters = AddModules.GetExportFilters(oForm);

                        DateTime dtFrom = DateTime.MinValue;
                        DateTime dtTo = DateTime.MinValue;

                        // 2. Safely convert the strings to DateTime
                        // (Only parse if the string is not empty)
                        if (!string.IsNullOrEmpty(filters.FromDate))
                        {
                            DateTime.TryParseExact(filters.FromDate, "yyyyMMdd",
                                System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None,
                                out dtFrom);
                        }

                        if (!string.IsNullOrEmpty(filters.ToDate))
                        {
                            DateTime.TryParseExact(filters.ToDate, "yyyyMMdd",
                                System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None,
                                out dtTo);
                        }

                        // 3. Call the exporter
                        ExcelExporter.ExportToExcel(
                            Company,
                            SBO_Application,
                            dtFrom,
                            dtTo,
                            filters.Sent,
                            filters.DocType
                        );
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.StatusBar.SetText(
                            "Excel export failed: " + ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Long,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    finally
                    {
                        // 5. ALWAYS reset the cursor back to default
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                    }
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ItemEvent error: {ex.Message}");
        }
    }


    static void SBO_Application_FormDataEvent(
      ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo,
      out bool BubbleEvent)
    {
        BubbleEvent = true;




        switch (BusinessObjectInfo.Type)
        {
            case "13":
                pathReportTemplates = Path.Combine(
                      AppContext.BaseDirectory,
                      "Reports",
                      "Invoice.rpt"
                  );
                AddModules.SaveDocument(SBO_Application, Company, ref BusinessObjectInfo, reportClass, "13", "A/R Invoice");
                break;
            case "14":
                pathReportTemplates = Path.Combine(
                     AppContext.BaseDirectory,
                     "Reports",
                     "CreditMemo.rpt"
                 );
                AddModules.SaveDocument(SBO_Application, Company, ref BusinessObjectInfo, reportClass, "14", "A/R Credit Memo");
                break;
        }

        if (BusinessObjectInfo.FormTypeEx == "BIR_FORM_1" &&
          BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE &&
          !BusinessObjectInfo.BeforeAction)
        {
            Console.WriteLine("BIR_FORM_1 closed by user.");
            // SAP keeps the form in memory; no cleanup needed here.
            // The next call to ShowAutoEmailForm will re-focus it.
        }


        Console.WriteLine("=================================");
    }

    static void SBO_Application_MenuEvent(
       ref SAPbouiCOM.MenuEvent pVal,
       out bool BubbleEvent)
    {
        BubbleEvent = true;

        try
        {
            if (pVal.BeforeAction == false)  // AFTER the click
            {
                if (pVal.MenuUID == "MNU_AUTOEMAIL")
                {
                    AddModules.ShowAutoEmailForm(SBO_Application, Company);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"MenuEvent error: {ex.Message}");
        }
    }

}
