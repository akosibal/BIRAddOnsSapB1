using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

public class ReportClass
{
    public byte[] GeneratePdf(
        string reportPath,
        Dictionary<string, object> parameters)
    {
        using var report = new ReportDocument();

        report.Load(reportPath);

        foreach (var parameter in parameters)
        {
            report.SetParameterValue(
                parameter.Key,
                parameter.Value);
        }

        using var stream = report.ExportToStream(
            ExportFormatType.PortableDocFormat);

        using var memoryStream = new MemoryStream();

        stream.CopyTo(memoryStream);

        return memoryStream.ToArray();
    }
}