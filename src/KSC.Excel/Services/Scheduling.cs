using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading;
using KSC.Excel.Engine;

namespace KSC.Excel.Services
{
    /// <summary>When a scheduled report runs (server local time).</summary>
    public sealed class ReportSchedule
    {
        readonly Func<DateTime, bool> _day;
        readonly TimeSpan _at;
        readonly TimeSpan? _every;

        ReportSchedule(Func<DateTime, bool> day, TimeSpan at, TimeSpan? every, string description) { _day = day; _at = at; _every = every; Description = description; }

        public string Description { get; private set; }

        public static ReportSchedule Daily(TimeSpan at) { return new ReportSchedule(d => true, at, null, "daily " + at); }

        public static ReportSchedule Weekly(DayOfWeek day, TimeSpan at) { return new ReportSchedule(d => d.DayOfWeek == day, at, null, "weekly " + day + " " + at); }

        /// <summary>Every month on this Gregorian day (clamped to the month length, so 31 means "last day" in short months).</summary>
        public static ReportSchedule Monthly(int day, TimeSpan at)
        {
            if (day < 1 || day > 31) throw new ArgumentOutOfRangeException("day");
            return new ReportSchedule(d => d.Day == Math.Min(day, DateTime.DaysInMonth(d.Year, d.Month)), at, null, "monthly day " + day + " " + at);
        }

        /// <summary>Every month on this Persian (Jalali) day (clamped to the month length: 31 = last day of the month).</summary>
        public static ReportSchedule MonthlyJalali(int day, TimeSpan at)
        {
            if (day < 1 || day > 31) throw new ArgumentOutOfRangeException("day");
            var pc = new PersianCalendar();
            return new ReportSchedule(d =>
            {
                int y, m, jd;
                if (!Jalali.TryParts(d, out y, out m, out jd)) return false;
                return jd == Math.Min(day, pc.GetDaysInMonth(y, m));
            }, at, null, "monthly (Jalali) day " + day + " " + at);
        }

        /// <summary>Every N minutes from the moment the scheduler starts (or the job is added).</summary>
        public static ReportSchedule EveryMinutes(int minutes)
        {
            if (minutes < 1) throw new ArgumentOutOfRangeException("minutes");
            return new ReportSchedule(d => true, TimeSpan.Zero, TimeSpan.FromMinutes(minutes), "every " + minutes + " min");
        }

        /// <summary>The first run time strictly after <paramref name="after"/>.</summary>
        public DateTime NextAfter(DateTime after)
        {
            if (_every.HasValue) return after + _every.Value;
            for (int i = 0; i <= 70; i++)
            {
                var date = after.Date.AddDays(i);
                if (!_day(date)) continue;
                var t = date + _at;
                if (t > after) return t;
            }
            return DateTime.MaxValue;
        }
    }

    /// <summary>A report to produce on a schedule and hand to one or more <see cref="IReportSink"/>s.</summary>
    public sealed class ReportJob
    {
        public string Name { get; set; }
        /// <summary>Name of a report registered with <c>options.AddReport</c>.</summary>
        public string Report { get; set; }
        public ReportFormat Format { get; set; } = ReportFormat.Pdf;
        public ReportSchedule Schedule { get; set; }
        /// <summary>Fixed report parameters.</summary>
        public IDictionary<string, object> Parameters { get; set; }
        /// <summary>Computes parameters at run time (e.g. "yesterday"); merged over <see cref="Parameters"/>.</summary>
        public Func<DateTime, IDictionary<string, object>> ParameterFactory { get; set; }
        /// <summary>E-mail recipients (needs <see cref="SmtpReportSink"/>).</summary>
        public IList<string> To { get; set; } = new List<string>();
        public string Subject { get; set; }
        public string Body { get; set; }
        /// <summary>Also save the file here (needs <see cref="FolderReportSink"/>).</summary>
        public string SaveToFolder { get; set; }

        internal DateTime Next;
        internal int Running;
    }

    /// <summary>Receives every produced file of a job.</summary>
    public interface IReportSink
    {
        void Deliver(ReportJob job, ReportFile file, string fileName, DateTime when);
    }

    /// <summary>Writes files into <see cref="ReportJob.SaveToFolder"/> (name + time stamp).</summary>
    public sealed class FolderReportSink : IReportSink
    {
        public void Deliver(ReportJob job, ReportFile file, string fileName, DateTime when)
        {
            if (string.IsNullOrWhiteSpace(job.SaveToFolder)) return;
            Directory.CreateDirectory(job.SaveToFolder);
            var name = Path.GetFileNameWithoutExtension(fileName) + "-" + when.ToString("yyyyMMdd-HHmm", CultureInfo.InvariantCulture) + "." + file.Extension;
            File.WriteAllBytes(Path.Combine(job.SaveToFolder, name), file.Bytes);
        }
    }

    public sealed class SmtpOptions
    {
        public string Host { get; set; }
        public int Port { get; set; } = 587;
        public bool EnableSsl { get; set; } = true;
        public string User { get; set; }
        public string Password { get; set; }
        public string From { get; set; }
    }

    /// <summary>Sends the file as an e-mail attachment to <see cref="ReportJob.To"/> through SMTP.</summary>
    public sealed class SmtpReportSink : IReportSink
    {
        readonly SmtpOptions _o;
        public SmtpReportSink(SmtpOptions options)
        {
            if (options == null || string.IsNullOrWhiteSpace(options.Host) || string.IsNullOrWhiteSpace(options.From)) throw new ArgumentException("SMTP Host and From are required.");
            _o = options;
        }

        public void Deliver(ReportJob job, ReportFile file, string fileName, DateTime when)
        {
            if (job.To == null || job.To.Count == 0) return;
            using (var msg = new MailMessage())
            using (var ms = new MemoryStream(file.Bytes))
            {
                msg.From = new MailAddress(_o.From);
                foreach (var to in job.To) msg.To.Add(to);
                msg.Subject = job.Subject ?? job.Name ?? job.Report;
                msg.SubjectEncoding = System.Text.Encoding.UTF8; msg.BodyEncoding = System.Text.Encoding.UTF8;
                msg.Body = job.Body ?? "";
                var att = new Attachment(ms, fileName, file.ContentType.Split(';')[0]);
                msg.Attachments.Add(att);
                using (var smtp = new SmtpClient(_o.Host, _o.Port) { EnableSsl = _o.EnableSsl })
                {
                    if (!string.IsNullOrEmpty(_o.User)) smtp.Credentials = new NetworkCredential(_o.User, _o.Password);
                    smtp.Send(msg);
                }
            }
        }
    }

    /// <summary>
    /// Produces registered reports on a schedule (daily / weekly / monthly / Jalali monthly / every N minutes) and delivers them
    /// by e-mail and/or to a folder. In-process timer: it runs while your application runs.
    /// </summary>
    public sealed class ReportScheduler : IDisposable
    {
        readonly KscExcelOptions _opt;
        readonly IExcelReportService _svc;
        readonly List<IReportSink> _sinks = new List<IReportSink>();
        readonly List<ReportJob> _jobs = new List<ReportJob>();
        readonly Func<DateTime> _clock;
        readonly object _lock = new object();
        Timer _timer;

        public ReportScheduler(KscExcelOptions options, IExcelReportService service, Func<DateTime> clock = null)
        {
            _opt = options; _svc = service; _clock = clock ?? (() => DateTime.Now);
            _sinks.Add(new FolderReportSink());
        }

        /// <summary>Called after every run: <c>(job, error)</c>; error is null on success. Use it for logging.</summary>
        public Action<ReportJob, Exception> OnRun { get; set; }

        public ReportScheduler AddSink(IReportSink sink) { _sinks.Add(sink); return this; }

        /// <summary>Convenience: e-mail through SMTP.</summary>
        public ReportScheduler UseSmtp(SmtpOptions smtp) { return AddSink(new SmtpReportSink(smtp)); }

        public IReadOnlyList<ReportJob> Jobs { get { lock (_lock) return _jobs.ToList(); } }

        public ReportScheduler Add(ReportJob job)
        {
            if (job == null) throw new ArgumentNullException("job");
            if (job.Schedule == null) throw new ArgumentException("A job needs a Schedule.");
            if (string.IsNullOrWhiteSpace(job.Report) || !_opt.Reports.ContainsKey(job.Report)) throw new ArgumentException("Report '" + job.Report + "' is not registered (options.AddReport).");
            job.Name = job.Name ?? job.Report;
            job.Next = job.Schedule.NextAfter(_clock());
            lock (_lock) _jobs.Add(job);
            return this;
        }

        public void Start(TimeSpan? tick = null)
        {
            if (_timer != null) return;
            var period = tick ?? TimeSpan.FromSeconds(30);
            _timer = new Timer(_ => { try { RunDueJobs(); } catch (Exception) { } }, null, period, period);
        }

        public void Stop() { var t = _timer; _timer = null; if (t != null) t.Dispose(); }

        public void Dispose() { Stop(); }

        /// <summary>Runs every job whose time has come (called by the timer; public so you can drive it yourself or test it). Returns the jobs started.</summary>
        public int RunDueJobs()
        {
            var now = _clock();
            List<ReportJob> due;
            lock (_lock) due = _jobs.Where(j => j.Next <= now).ToList();
            foreach (var j in due)
            {
                j.Next = j.Schedule.NextAfter(now);        // schedule first: a slow or failing run never causes a burst of catch-up runs
                RunJob(j, now);
            }
            return due.Count;
        }

        /// <summary>Runs one job immediately (ignores the schedule).</summary>
        public void RunNow(string jobName)
        {
            ReportJob job;
            lock (_lock) job = _jobs.FirstOrDefault(j => string.Equals(j.Name, jobName, StringComparison.OrdinalIgnoreCase));
            if (job == null) throw new ArgumentException("Unknown job '" + jobName + "'.");
            RunJob(job, _clock());
        }

        void RunJob(ReportJob job, DateTime now)
        {
            if (Interlocked.CompareExchange(ref job.Running, 1, 0) != 0) return;      // still running from the previous tick
            Exception error = null; var sw = System.Diagnostics.Stopwatch.StartNew(); string paramText = "";
            try
            {
                var def = _opt.Reports[job.Report];
                var tpl = _svc.GetTemplate(def.Template ?? def.Name);
                var pars = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                if (job.Parameters != null) foreach (var kv in job.Parameters) pars[kv.Key] = kv.Value;
                if (job.ParameterFactory != null) { var extra = job.ParameterFactory(now); if (extra != null) foreach (var kv in extra) pars[kv.Key] = kv.Value; }
                paramText = ReportSecurity.Params(pars);
                var data = def.Load(pars, null);                 // scheduled runs have no user: role rules do not apply, the job itself is the authority
                var file = _svc.Export(_opt, tpl, data, pars, job.Format);
                var fileName = ReportExport.FileName(job.Name, file.Extension);
                foreach (var sink in _sinks) sink.Deliver(job, file, fileName, now);
            }
            catch (Exception ex) { error = ex; }
            finally { Interlocked.Exchange(ref job.Running, 0); }
            ReportSecurity.Write(_opt, new ReportAuditEntry { Time = DateTime.Now, User = "scheduler", Report = job.Report, Action = "scheduled:" + job.Format.ToString().ToLowerInvariant(), Status = error == null ? ReportAuditStatus.Ok : ReportAuditStatus.Error, Parameters = paramText, DurationMs = sw.ElapsedMilliseconds, Detail = error == null ? job.Name : error.GetType().Name });
            var cb = OnRun;
            if (cb != null) { try { cb(job, error); } catch (Exception) { } }
        }
    }
}
