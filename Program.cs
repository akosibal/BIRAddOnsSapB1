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

        var conn = new Conn();

        conn.Connect(args[0]);

        Console.WriteLine("Add-on started.");

        Console.ReadLine();
    }
}
