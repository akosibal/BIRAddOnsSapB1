using System;
using System.Collections.Generic;
using System.IO;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

public class ReportClass
{
    // ============================================================
    // Database Configuration
    // ============================================================


    private const string OutputPath = @"D:\System\BIRAddOnsSapB1\CrystalReport\Output\Invoice.pdf";

    string _dbServer { get; set; } = "192.168.5.22";     // e.g. "SAPSRV" or "SAPSRV,1433"
    string _dbName { get; set; } = "TEST_MDTI_20260732";     // e.g. "SBODEMOUS"
    string _dbUser { get; set; } = "sa";
    string _dbPassword { get; set; } = "1q2w#E$R";

    public ReportClass(
        string dbServer,
        string dbName,
        string dbUser,
        string dbPassword)
    {
        _dbServer = dbServer;
        _dbName = dbName;
        _dbUser = dbUser;
        _dbPassword = dbPassword;
    }

    // ============================================================
    // Generate PDF
    // ============================================================

    public byte[] GeneratePdf(
        string reportPath,
        Dictionary<string, object> parameters)
    {
        if (string.IsNullOrWhiteSpace(reportPath))
            throw new ArgumentException(
                "Report path is required.",
                nameof(reportPath));

        if (!File.Exists(reportPath))
            throw new FileNotFoundException(
                $"Crystal Report not found: {reportPath}",
                reportPath);

        ReportDocument report = null;

        try
        {
            report = new ReportDocument();

            // ----------------------------------------------------
            // Load report
            // ----------------------------------------------------

            report.Load(reportPath);

            // ----------------------------------------------------
            // Apply database logon
            // ----------------------------------------------------

            ApplyLogon(report);

            // ----------------------------------------------------
            // Apply MAIN report parameters only
            // ----------------------------------------------------

            ApplyMainParameters(
                report,
                parameters ?? new Dictionary<string, object>());

            // ----------------------------------------------------
            // Export PDF to memory
            // ----------------------------------------------------

            report.ExportToDisk(ExportFormatType.PortableDocFormat, OutputPath);

            using Stream stream =
                report.ExportToStream(
                    ExportFormatType.PortableDocFormat);

            using MemoryStream memoryStream =
                new MemoryStream();

            stream.CopyTo(memoryStream);

            return memoryStream.ToArray();
        }
        finally
        {
            if (report != null)
            {
                try
                {
                    report.Close();
                }
                catch
                {
                }

                try
                {
                    report.Dispose();
                }
                catch
                {
                }
            }
        }
    }

    // ============================================================
    // Apply Database Logon
    // ============================================================

    private void ApplyLogon(ReportDocument report)
    {
        var logonInfo = new ConnectionInfo
        {
            ServerName = _dbServer,
            DatabaseName = _dbName,
            UserID = _dbUser,
            Password = _dbPassword,
            IntegratedSecurity = false
        };

        // --------------------------------------------------------
        // Main report tables
        // --------------------------------------------------------

        foreach (Table table in report.Database.Tables)
        {
            SetTableLogon(table, logonInfo);
        }

        // --------------------------------------------------------
        // Subreports
        //
        // IMPORTANT:
        // Only enumerate Subreports from the MAIN report.
        // Do NOT recursively call subreport.Subreports because
        // Crystal throws:
        //
        // NotSupportedException:
        // Not supported within subreports.
        // --------------------------------------------------------

        foreach (ReportDocument subreport in report.Subreports)
        {
            ApplySubreportLogon(
                subreport,
                logonInfo);
        }
    }

    private void ApplySubreportLogon(
        ReportDocument subreport,
        ConnectionInfo logonInfo)
    {
        foreach (Table table in subreport.Database.Tables)
        {
            SetTableLogon(
                table,
                logonInfo);
        }
    }

    private void SetTableLogon(
        Table table,
        ConnectionInfo logonInfo)
    {
        TableLogOnInfo tableLogon =
            table.LogOnInfo;

        tableLogon.ConnectionInfo =
            logonInfo;

        table.ApplyLogOnInfo(
            tableLogon);

        // Force Crystal to refresh the table location.
        table.Location = table.Location;
    }

    // ============================================================
    // Main Report Parameters
    // ============================================================

    private void ApplyMainParameters(
        ReportDocument report,
        Dictionary<string, object> parameters)
    {
        foreach (ParameterFieldDefinition param
                 in report.DataDefinition.ParameterFields)
        {
            // Linked parameters are handled by Crystal through
            // the subreport links.
            bool linked = SafeIsLinked(param);

            if (linked)
                continue;

            if (!parameters.TryGetValue(
                    param.Name,
                    out object value))
            {
                continue;
            }

            SetParameter(
                report,
                param,
                value);
        }
    }

    // ============================================================
    // Set Parameter
    // ============================================================

    private void SetParameter(
        ReportDocument report,
        ParameterFieldDefinition param,
        object rawValue)
    {
        var dv = new ParameterDiscreteValue();

        switch (param.ValueType)
        {
            case CrystalDecisions.Shared.FieldValueType.NumberField:
            case CrystalDecisions.Shared.FieldValueType.Int32sField:
            case CrystalDecisions.Shared.FieldValueType.Int16sField:
            case CrystalDecisions.Shared.FieldValueType.Int32uField:
                dv.Value = Convert.ToInt64(rawValue);
                break;
            case CrystalDecisions.Shared.FieldValueType.CurrencyField:
                dv.Value = Convert.ToDecimal(rawValue);
                break;
            case CrystalDecisions.Shared.FieldValueType.DateField:
            case CrystalDecisions.Shared.FieldValueType.DateTimeField:
                dv.Value = Convert.ToDateTime(rawValue);
                break;
            case CrystalDecisions.Shared.FieldValueType.BooleanField:
                dv.Value = Convert.ToBoolean(rawValue);
                break;
            default:
                dv.Value = rawValue?.ToString() ?? string.Empty;
                break;
        }

        report.ParameterFields[param.Name].CurrentValues.Clear();
        report.ParameterFields[param.Name].CurrentValues.Add(dv);
    }

    // ============================================================
    // Safe IsLinked
    // ============================================================

    private bool SafeIsLinked(
        ParameterFieldDefinition param)
    {
        try
        {
            return param.IsLinked();
        }
        catch
        {
            return false;
        }
    }
}