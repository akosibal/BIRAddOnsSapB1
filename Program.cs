using SAPbouiCOM;
using SAPbobsCOM;

/* Console.WriteLine("Hello, World!");

Console.WriteLine("SAP Business One Add-on");

Console.WriteLine("UI API: " + typeof(SboGuiApi).FullName);

Console.WriteLine("DI API: " + typeof(SAPbobsCOM.Company).FullName);

Console.ReadLine();
 */
class Program
{
    private static SAPbouiCOM.Application SBO_Application;
    private static SAPbobsCOM.Company Company;

    static Conn conn = new Conn();

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
            var conn = new Conn();

            conn.Connect(args[0]);

            // Reuse the existing connections
            SBO_Application = conn.SBO_Application;
            Company = conn.company;

            Console.WriteLine("Add-on started.");

            SBO_Application.FormDataEvent +=
                SBO_Application_FormDataEvent;

            Console.WriteLine("FormDataEvent registered.");
            Console.WriteLine("Waiting for SAP Business One events...");

            System.Threading.Thread.Sleep(
                System.Threading.Timeout.Infinite
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
    }
    static void LoadForm(Conn conn)
    {
        Console.WriteLine("Loading form:");

        string formPath = Path.Combine(
            AppContext.BaseDirectory,
            "Forms",
            "Form1.b1f"
        );

        Console.WriteLine(formPath);

        if (!File.Exists(formPath))
        {
            throw new FileNotFoundException(
                "Form1.b1f was not found.",
                formPath
            );
        }

        Console.WriteLine("Form file found.");

        string xml = File.ReadAllText(formPath);

        Console.WriteLine("XML loaded.");
        Console.WriteLine($"XML length: {xml.Length}");

        Console.WriteLine("Checking SBO_Application...");

        if (conn.SBO_Application == null)
        {
            throw new Exception(
                "SBO_Application is NULL. The SAP Business One UI API connection was not initialized."
            );
        }

        Console.WriteLine("SBO_Application OK.");

        Console.WriteLine("Calling LoadBatchActions...");

        conn.SBO_Application.LoadBatchActions(ref xml);

        Console.WriteLine("LoadBatchActions completed.");

        Console.WriteLine("Getting BIR_FORM_1...");

        SAPbouiCOM.Form form =
            conn.SBO_Application.Forms.Item("BIR_FORM_1");

        if (form == null)
        {
            throw new Exception(
                "BIR_FORM_1 was not found after loading Form1.b1f."
            );
        }

        Console.WriteLine("BIR_FORM_1 found.");

        form.Visible = true;

        Console.WriteLine("BIR_FORM_1 is now visible.");
    }


    static void SBO_Application_FormDataEvent(
     ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo,
     out bool BubbleEvent)
    {
        BubbleEvent = true;

        // Sales Invoice
        if (BusinessObjectInfo.Type != "13")
            return;

        // ADD only
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

        try
        {
            SAPbouiCOM.Form oForm =
                SBO_Application.Forms.Item(
                    BusinessObjectInfo.FormUID
                );

            Console.WriteLine($"Form Type: {oForm.TypeEx}");

            SAPbouiCOM.DBDataSource oDBDataSource =
                oForm.DataSources.DBDataSources.Item("OINV");

            string docEntry =
                oDBDataSource.GetValue("DocEntry", 0).Trim();

            Console.WriteLine($"DocEntry: {docEntry}");

            if (string.IsNullOrEmpty(docEntry))
            {
                Console.WriteLine("DocEntry is empty.");
                return;
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
        }

        Console.WriteLine("===================+==============");
    }

}
