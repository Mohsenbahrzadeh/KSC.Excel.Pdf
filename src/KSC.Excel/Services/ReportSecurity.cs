using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace KSC.Excel.Services
{
    /// <summary>What a report's data function receives: the typed parameters and WHO is asking (null = a scheduled job / the system).</summary>
    public sealed class ReportContext
    {
        public string ReportName { get; set; }
        public IDictionary<string, object> Parameters { get; set; }
        /// <summary>The signed-in user; use it to filter data per user (row-level security). Null for scheduled jobs.</summary>
        public IPrincipal User { get; set; }
    }

    public enum ReportAuditStatus { Ok, Denied, Error }

    /// <summary>One line of the audit trail.</summary>
    public sealed class ReportAuditEntry
    {
        public DateTime Time { get; set; }
        /// <summary>User name, or "scheduler" / "anonymous".</summary>
        public string User { get; set; }
        public string Report { get; set; }
        /// <summary>view, pdf, xlsx, docx, csv, html, png or scheduled:&lt;format&gt;.</summary>
        public string Action { get; set; }
        public ReportAuditStatus Status { get; set; }
        /// <summary>Parameter values (secrets such as password/token/key are left out; long values are cut).</summary>
        public string Parameters { get; set; }
        public long DurationMs { get; set; }
        /// <summary>Reason for Denied / the error type for Error (never the message of the exception).</summary>
        public string Detail { get; set; }
    }

    /// <summary>Receives an entry for every report run, denial and failure.</summary>
    public interface IReportAudit
    {
        void Log(ReportAuditEntry entry);
    }

    /// <summary>Calls your delegate (write to your log, a database...).</summary>
    public sealed class DelegateReportAudit : IReportAudit
    {
        readonly Action<ReportAuditEntry> _log;
        public DelegateReportAudit(Action<ReportAuditEntry> log) { _log = log; }
        public void Log(ReportAuditEntry entry) { _log(entry); }
    }

    /// <summary>Append-only JSON-lines files, one per day: <c>audit-20251019.jsonl</c>.</summary>
    public sealed class FileReportAudit : IReportAudit
    {
        readonly string _dir;
        readonly object _lock = new object();
        public FileReportAudit(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("directory");
            _dir = Path.GetFullPath(directory);
        }

        public void Log(ReportAuditEntry e)
        {
            var line = JsonConvert.SerializeObject(new { time = e.Time.ToString("o", CultureInfo.InvariantCulture), user = e.User, report = e.Report, action = e.Action, status = e.Status.ToString(), parameters = e.Parameters, ms = e.DurationMs, detail = e.Detail });
            lock (_lock)
            {
                Directory.CreateDirectory(_dir);
                File.AppendAllText(Path.Combine(_dir, "audit-" + e.Time.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".jsonl"), line + "\n", new UTF8Encoding(false));
            }
        }
    }

    static class ReportSecurity
    {
        static readonly Regex Secret = new Regex("pass|secret|token|key|pwd|auth", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>401 = sign in, 403 = signed in but not allowed, 0 = allowed.</summary>
        public static int Check(ReportDef def, IPrincipal user)
        {
            bool restricted = (def.Roles != null && def.Roles.Count > 0) || def.Authorize != null;
            if (!restricted) return 0;
            if (user == null || user.Identity == null || !user.Identity.IsAuthenticated) return 401;
            if (def.Roles != null && def.Roles.Count > 0 && !def.Roles.Any(r => user.IsInRole(r))) return 403;
            if (def.Authorize != null && !def.Authorize(user)) return 403;
            return 0;
        }

        public static string UserName(IPrincipal user)
            => user != null && user.Identity != null && user.Identity.IsAuthenticated ? user.Identity.Name ?? "(unnamed)" : "anonymous";

        public static string Params(IDictionary<string, object> p)
        {
            if (p == null || p.Count == 0) return "";
            var parts = new List<string>();
            foreach (var kv in p.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
            {
                if (Secret.IsMatch(kv.Key)) { parts.Add(kv.Key + "=***"); continue; }
                var v = kv.Value is DateTime ? ((DateTime)kv.Value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : Convert.ToString(kv.Value, CultureInfo.InvariantCulture) ?? "";
                v = v.Replace('\r', ' ').Replace('\n', ' ');
                parts.Add(kv.Key + "=" + (v.Length > 100 ? v.Substring(0, 100) + "..." : v));
            }
            return string.Join("; ", parts);
        }

        public static void Write(KscExcelOptions o, ReportAuditEntry e)
        {
            var a = o == null ? null : o.Audit;
            if (a == null) return;
            try { a.Log(e); } catch (Exception) { /* an audit failure must never break reporting */ }
        }
    }
}
