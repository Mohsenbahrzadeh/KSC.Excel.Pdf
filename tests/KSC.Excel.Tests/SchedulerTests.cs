using System.Net;
using System.Net.Sockets;
using System.Text;
using KSC.Excel.Model;
using KSC.Excel.Services;
using Xunit;

namespace KSC.Excel.Tests;

public class SchedulerTests
{
    static readonly TimeSpan Eight = TimeSpan.FromHours(8);

    // ------------------------------------------------------------------ schedule arithmetic
    [Fact]
    public void Daily_and_weekly_pick_the_next_moment_strictly_after()
    {
        var d = ReportSchedule.Daily(Eight);
        Assert.Equal(new DateTime(2025, 3, 10, 8, 0, 0), d.NextAfter(new DateTime(2025, 3, 10, 7, 59, 0)));
        Assert.Equal(new DateTime(2025, 3, 11, 8, 0, 0), d.NextAfter(new DateTime(2025, 3, 10, 8, 0, 0)));      // exactly now: tomorrow
        var w = ReportSchedule.Weekly(DayOfWeek.Saturday, Eight);
        Assert.Equal(new DateTime(2025, 3, 15, 8, 0, 0), w.NextAfter(new DateTime(2025, 3, 10, 12, 0, 0)));        // 2025-03-15 is a Saturday
    }

    [Fact]
    public void Monthly_clamps_to_short_months()
    {
        var m = ReportSchedule.Monthly(31, Eight);
        Assert.Equal(new DateTime(2025, 2, 28, 8, 0, 0), m.NextAfter(new DateTime(2025, 2, 1)));
        Assert.Equal(new DateTime(2025, 4, 30, 8, 0, 0), m.NextAfter(new DateTime(2025, 4, 1)));
        Assert.Equal(new DateTime(2024, 2, 29, 8, 0, 0), m.NextAfter(new DateTime(2024, 2, 1)));                  // leap year
        Assert.Equal(new DateTime(2025, 5, 5, 8, 0, 0), ReportSchedule.Monthly(5, Eight).NextAfter(new DateTime(2025, 4, 6)));
    }

    [Fact]
    public void Jalali_monthly_follows_the_persian_calendar()
    {
        // 1404/01/01 = 2025-03-21 ; day 15 of Farvardin = 2025-04-04 ; Esfand 1403 has 30 days (leap) -> "31" = last day = 1403/12/30 = 2025-03-20
        Assert.Equal(new DateTime(2025, 4, 4, 8, 0, 0), ReportSchedule.MonthlyJalali(15, Eight).NextAfter(new DateTime(2025, 3, 25)));
        Assert.Equal(new DateTime(2025, 3, 20, 8, 0, 0), ReportSchedule.MonthlyJalali(31, Eight).NextAfter(new DateTime(2025, 3, 10)));
        // first six months have 31 days: 1404/02/31 = 2025-05-21
        Assert.Equal(new DateTime(2025, 5, 21, 8, 0, 0), ReportSchedule.MonthlyJalali(31, Eight).NextAfter(new DateTime(2025, 5, 1)));
    }

    [Fact]
    public void Every_minutes_and_argument_validation()
    {
        Assert.Equal(new DateTime(2025, 1, 1, 0, 15, 0), ReportSchedule.EveryMinutes(15).NextAfter(new DateTime(2025, 1, 1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => ReportSchedule.EveryMinutes(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => ReportSchedule.Monthly(32, Eight));
    }

    // ------------------------------------------------------------------ scheduler
    sealed class FakeSink : IReportSink
    {
        public readonly List<(string job, string file, int bytes, DateTime when)> Delivered = new();
        public void Deliver(ReportJob job, ReportFile f, string fileName, DateTime when) => Delivered.Add((job.Name, fileName, f.Bytes.Length, when));
    }

    static (ReportScheduler sch, FakeSink sink, Func<DateTime> now, Action<DateTime> set, List<IDictionary<string, object>> seen) Setup()
    {
        var store = new MemoryTemplateStore();
        var b = new TemplateBuilder("Daily", "d"); b.Sheet("S", s => s.Cell("A1", "{{$p.Day|fa}}").Cell("B1", "{{Name}}"));
        var t = b.Build(); t.Parameters.Add(new ParameterDef { Name = "Day", Type = "date" });
        store.Save(t);
        var opt = new KscExcelOptions();
        var seen = new List<IDictionary<string, object>>();
        opt.AddReport("Daily", p => { seen.Add(p); return new { Name = "x" }; });
        var clock = new DateTime(2025, 3, 10, 7, 0, 0);
        var sch = new ReportScheduler(opt, new ExcelReportService(opt, store), () => clock);
        var sink = new FakeSink(); sch.AddSink(sink);
        return (sch, sink, () => clock, v => clock = v, seen);
    }

    [Fact]
    public void Due_jobs_run_once_deliver_and_reschedule_without_catching_up_in_a_burst()
    {
        var (sch, sink, now, set, seen) = Setup();
        sch.Add(new ReportJob { Name = "Morning", Report = "Daily", Format = ReportFormat.Csv, Schedule = ReportSchedule.Daily(Eight), ParameterFactory = t => new Dictionary<string, object> { ["Day"] = t.Date.AddDays(-1) } });
        Assert.Equal(0, sch.RunDueJobs());                                     // 07:00: not yet
        set(new DateTime(2025, 3, 10, 8, 0, 5));
        Assert.Equal(1, sch.RunDueJobs());
        Assert.Single(sink.Delivered); Assert.Equal("Morning.csv", sink.Delivered[0].file);
        Assert.Equal(new DateTime(2025, 3, 9), seen[0]["Day"]);                // parameter factory: yesterday
        Assert.Equal(0, sch.RunDueJobs());                                     // same tick again: nothing
        set(new DateTime(2025, 3, 14, 9, 0, 0));                               // server was down for days: one run, not four
        Assert.Equal(1, sch.RunDueJobs()); Assert.Equal(2, sink.Delivered.Count);
        set(new DateTime(2025, 3, 15, 8, 0, 1)); Assert.Equal(1, sch.RunDueJobs());
    }

    [Fact]
    public void Failures_are_reported_and_do_not_stop_other_jobs()
    {
        var (sch, sink, now, set, seen) = Setup();
        var errors = new List<(string, Exception)>();
        sch.OnRun = (j, e) => errors.Add((j.Name, e));
        var opt2 = typeof(ReportScheduler);
        sch.Add(new ReportJob { Name = "Good", Report = "Daily", Schedule = ReportSchedule.EveryMinutes(1), Format = ReportFormat.Csv });
        sch.Add(new ReportJob { Name = "Bad", Report = "Daily", Schedule = ReportSchedule.EveryMinutes(1), Format = ReportFormat.Pdf });   // no PDF package: fails
        set(now().AddMinutes(2));
        Assert.Equal(2, sch.RunDueJobs());
        Assert.Single(sink.Delivered);
        Assert.Null(errors.Single(e => e.Item1 == "Good").Item2);
        Assert.Contains("UsePdf", errors.Single(e => e.Item1 == "Bad").Item2.Message);
    }

    [Fact]
    public void Run_now_validation_and_folder_delivery()
    {
        var (sch, sink, now, set, seen) = Setup();
        var dir = Path.Combine(Path.GetTempPath(), "kxl-sched-" + Guid.NewGuid().ToString("N"));
        try
        {
            sch.Add(new ReportJob { Name = "Saved", Report = "Daily", Format = ReportFormat.Docx, Schedule = ReportSchedule.Daily(Eight), SaveToFolder = dir });
            sch.RunNow("saved");
            var files = Directory.GetFiles(dir);
            Assert.Single(files); Assert.Matches(@"Saved-20250310-0700\.docx$", files[0]);
            Assert.True(new FileInfo(files[0]).Length > 500);
            Assert.Throws<ArgumentException>(() => sch.RunNow("nope"));
            Assert.Throws<ArgumentException>(() => sch.Add(new ReportJob { Report = "Unregistered", Schedule = ReportSchedule.Daily(Eight) }));
            Assert.Throws<ArgumentException>(() => sch.Add(new ReportJob { Report = "Daily" }));
        }
        finally { if (Directory.Exists(dir)) Directory.Delete(dir, true); }
    }

    // ------------------------------------------------------------------ SMTP (a tiny local server)
    [Fact]
    public void Smtp_sink_sends_the_file_as_an_attachment()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start();
        int port = ((IPEndPoint)listener.LocalEndpoint).Port;
        string received = null;
        var server = Task.Run(() =>
        {
            using var client = listener.AcceptTcpClient(); using var stream = client.GetStream();
            var reader = new StreamReader(stream, Encoding.ASCII); var writer = new StreamWriter(stream, Encoding.ASCII) { NewLine = "\r\n", AutoFlush = true };
            writer.WriteLine("220 test");
            var data = new StringBuilder(); bool inData = false;
            while (true)
            {
                var line = reader.ReadLine(); if (line == null) break;
                if (inData) { if (line == ".") { inData = false; writer.WriteLine("250 queued"); received = data.ToString(); } else data.AppendLine(line); continue; }
                var cmd = line.ToUpperInvariant();
                if (cmd.StartsWith("EHLO") || cmd.StartsWith("HELO")) writer.WriteLine("250 hello");
                else if (cmd.StartsWith("DATA")) { writer.WriteLine("354 go"); inData = true; }
                else if (cmd.StartsWith("QUIT")) { writer.WriteLine("221 bye"); break; }
                else writer.WriteLine("250 ok");
            }
        });
        try
        {
            var sink = new SmtpReportSink(new SmtpOptions { Host = "127.0.0.1", Port = port, EnableSsl = false, From = "reports@example.com" });
            var job = new ReportJob { Name = "Weekly", Report = "x", To = { "boss@example.com", "cfo@example.com" }, Subject = "گزارش هفتگی", Body = "attached" };
            sink.Deliver(job, new ReportFile { Bytes = Encoding.UTF8.GetBytes("a,b\r\n1,2\r\n"), ContentType = "text/csv; charset=utf-8", Extension = "csv" }, "Weekly.csv", DateTime.Now);
            Assert.Same(server, Task.WhenAny(server, Task.Delay(5000)).GetAwaiter().GetResult());
            Assert.True(received != null, "nothing received"); Assert.Contains("boss@example.com", received); Assert.Contains("cfo@example.com", received);
            Assert.Contains("name=Weekly.csv", received.Replace("\"", ""));
            Assert.Contains(Convert.ToBase64String(Encoding.UTF8.GetBytes("a,b\r\n1,2\r\n")), received.Replace("\r\n", ""));
        }
        finally { listener.Stop(); }
        Assert.Throws<ArgumentException>(() => new SmtpReportSink(new SmtpOptions()));
    }
}
