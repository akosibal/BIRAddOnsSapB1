
using SAPbouiCOM;
using SAPbobsCOM;

public class Conn
{
    private SboGuiApi sboGuiApi;
    private Application application;
    private SAPbobsCOM.Company company;

    public void Connect(string connectionString)
    {
        try
        {
            // Create UI API connection
            sboGuiApi = new SboGuiApi();

            // Connect to SAP Business One
            sboGuiApi.Connect(connectionString);

            // Get SAP Business One application
            application = sboGuiApi.GetApplication(-1);

            Console.WriteLine("Connected to SAP Business One UI API.");

            // Get DI API company from UI API
            company = (SAPbobsCOM.Company)application.Company.GetDICompany();

            Console.WriteLine(
                $"Connected to SAP Business One DI API: {company.CompanyName}"
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine("SAP Business One connection failed.");
            Console.WriteLine(ex);
        }
    }
}
