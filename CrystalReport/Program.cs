using System;
using System.Collections.Generic;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace GenerateReport
{
    internal class Program
    {
        // ===== CONFIG =====
        private const string ReportPath = @"D:\System\BIRAddOnsSapB1\CrystalReport\Reports\Invoice.rpt";
        private const string OutputPath = @"D:\System\BIRAddOnsSapB1\CrystalReport\Output\Invoice.pdf";

        // Database logon (adjust to your SAP B1 / SQL Server)
        private const string DbServer   = "YOUR_SQL_SERVER";     // e.g. "SAPSRV" or "SAPSRV,1433"
        private const string DbName     = "YOUR_COMPANY_DB";     // e.g. "SBODEMOUS"
        private const string DbUser     = "sa";
        private const string DbPassword = "your_password";

        // Parameter values to inject (name -> value). Case-insensitive.
        private static readonly Dictionary<string, object> ParameterValues =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
        {
            { "@dockey",    "139320" },
            { "@objectid",  "17"     },
            { "@usercode",  "bal"    },
        };
        // ==================

        static void Main(string[] args)
        {
            Console.WriteLine("Crystal Report Runner");

            ReportDocument report = null;
            try
            {
                report = new ReportDocument();
                report.Load(ReportPath);

                // 1) Apply DB logon to main report + all subreports
                ApplyLogon(report);

                // 2) Set parameters on main report + all subreports (skip linked ones)
                ApplyParameters(report, isMainReport: true);

                // 3) Export
                report.ExportToDisk(ExportFormatType.PortableDocFormat, OutputPath);
                Console.WriteLine($"PDF created: {OutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:");
                Console.WriteLine(ex);
            }
            finally
            {
                if (report != null)
                {
                    try { report.Close(); } catch { }
                    try { report.Dispose(); } catch { }
                }
            }
        }

        // ---------------------------------------------------------------
        // Apply DB logon to main report and recursively to subreports
        // ---------------------------------------------------------------
        private static void ApplyLogon(ReportDocument report)
        {
            var logonInfo = new ConnectionInfo
            {
                ServerName   = DbServer,
                DatabaseName = DbName,
                UserID       = DbUser,
                Password     = DbPassword,
                IntegratedSecurity = false
            };

            // Main report tables
            foreach (Table table in report.Database.Tables)
                SetTableLogon(table, logonInfo);

            // Recurse into subreports
            foreach (ReportDocument sub in report.Subreports)
                ApplyLogon(sub);
        }

        private static void SetTableLogon(Table table, ConnectionInfo logonInfo)
        {
            TableLogOnInfo tableLogon = table.LogOnInfo;
            tableLogon.ConnectionInfo = logonInfo;
            table.ApplyLogOnInfo(tableLogon);

            // Some reports need Location reset to the correct DB
            table.Location = table.Location; // no-op but triggers refresh in some versions
        }

        // ---------------------------------------------------------------
        // Recursively set parameter values (skip linked parameters)
        // ---------------------------------------------------------------
        private static void ApplyParameters(ReportDocument report, bool isMainReport)
        {
            string label = isMainReport ? "Main report" : $"Subreport [{report.Name}]";
            Console.WriteLine($"\n--- {label}: scanning parameters ---");

            foreach (ParameterFieldDefinition param in report.DataDefinition.ParameterFields)
            {
                bool linked = SafeIsLinked(param);
                Console.WriteLine($"  {param.Name}  Linked={linked}  ReportName='{param.ReportName}'");

                // Linked parameters are fed automatically from the parent — do NOT set them
                if (linked) continue;

                if (ParameterValues.TryGetValue(param.Name, out object value))
                {
                    try
                    {
                        SetParameter(report, param, value);
                        Console.WriteLine($"    -> set {param.Name} = {value}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"    !! Failed to set {param.Name}: {ex.Message}");
                    }
                }
            }

            // Recurse into subreports
            foreach (ReportDocument sub in report.Subreports)
                ApplyParameters(sub, isMainReport: false);
        }

        // ---------------------------------------------------------------
        // Set a parameter respecting its declared type
        // ---------------------------------------------------------------
        private static void SetParameter(ReportDocument report, ParameterFieldDefinition param, object rawValue)
        {
            var dv = new ParameterDiscreteValue();

            switch (param.ValueType)
            {
                case CrystalDecisions.Shared.FieldValueType.NumberField:
                case CrystalDecisions.Shared.FieldValueType.Int32sField:
                case CrystalDecisions.Shared.FieldValueType.Int16sField:
                case CrystalDecisions.Shared.FieldValueType.Int64sField:
                    dv.Value = Convert.ToInt64(rawValue);
                    break;
                case CrystalDecisions.Shared.FieldValueType.DecimalField:
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

        // ---------------------------------------------------------------
        // IsLinked can throw on some parameter types — guard it
        // ---------------------------------------------------------------
        private static bool SafeIsLinked(ParameterFieldDefinition param)
        {
            try { return param.IsLinked(); }
            catch { return false; }
        }
    }
}