using System.Security.Principal;
using System.Text;
using KSC.Excel.Designer;
using KSC.Excel.Model;
using KSC.Excel.Services;
using Xunit;

namespace KSC.Excel.Tests;

public class ReportSecurityTests
{
    static IPrincipal User(string name, params string[] roles) => new GenericPrincipal(new GenericIdentity(name, "test"), roles);
    static readonly IPrincipal Anonymous = new GenericPrincipal(new GenericIdentity(""), new string[0]);

    sealed class MemAudit : IReportAudit { public readonly List<ReportAuditEntry> Entries = new(); public void Log(ReportAuditEntry e) => Entries.Add(e); }

    static (ReportServer srv, MemAudit audit, KscExcelOptions opt, ExcelReportService svc) Setup()
    {
        var store = new MemoryTemplateStore();
        var b = new TemplateBuilder("Orders", "Orders"); b.Sheet("S", s => { s.Cell("A1", "{{Owner}}").Cell("B1", "{{Total}}"); s.Band("R", "Rows", 1, 1); });
        var t = b.Build(); t.Parameters.Add(new ParameterDef { Name = "Branch", Type = "text" }); t.Parameters.Add(new ParameterDef { Name = "ApiKey", Type = "text" });
        store.Save(t);
        var all = new[] { new { Owner = "ali", Total = 10 }, new { Owner = "sara", Total = 20 }, new { Owner = "ali", Total = 5 } };
        var opt = new KscExcelOptions(); var audit = new MemAudit(); opt.Audit = audit;
        opt.AddReport("Public", p => new { Rows = all }, template: "Orders");
        opt.AddReport("Finance", p => new { Rows = all }, template: "Orders", roles: new[] { "Finance", "Admin" });
        opt.AddReport("Corp", p => new { Rows = all }, template: "Orders", authorize: u => u.Identity.Name.EndsWith("@corp.com"));
        opt.AddContextReport("Mine", c => new { Rows = all.Where(x => x.Owner == c.User.Identity.Name).ToArray() }, template: "Orders", roles: new[] { "Sales" });
        var svc = new ExcelReportService(opt, store);
        return (new ReportServer(opt, svc), audit, opt, svc);
    }

    static ReportRequest Req(string name, IPrincipal user, params string[] kv)
    {
        var r = new ReportRequest { Name = name, User = user };
        for (int i = 0; i + 1 < kv.Length; i += 2) r.Query[kv[i]] = kv[i + 1];
        return r;
    }

    [Fact]
    public void Role_rules_return_401_when_anonymous_403_when_not_allowed_200_when_allowed()
    {
        var (srv, _, _, _) = Setup();
        Assert.Equal(200, srv.Handle(Req("Public", null)).Status);
        Assert.Equal(401, srv.Handle(Req("Finance", null)).Status);
        Assert.Equal(401, srv.Handle(Req("Finance", Anonymous)).Status);
        Assert.Equal(403, srv.Handle(Req("Finance", User("bob", "Sales"))).Status);
        Assert.Equal(200, srv.Handle(Req("Finance", User("amir", "Finance"))).Status);
        Assert.Equal(200, srv.Handle(Req("Finance", User("root", "Admin"))).Status);
        Assert.Equal(403, srv.Handle(Req("Finance", User("amir", "Finance"), "format", "xlsx")).Status == 200 ? 403 : 0);    // allowed users may export too
    }

    [Fact]
    public void Custom_authorize_rule_and_catalog_filtering()
    {
        var (srv, _, _, _) = Setup();
        Assert.Equal(403, srv.Handle(Req("Corp", User("x@gmail.com"))).Status);
        Assert.Equal(200, srv.Handle(Req("Corp", User("x@corp.com"))).Status);
        var cat = Encoding.UTF8.GetString(srv.Handle(Req(null, User("x@corp.com", "Sales"))).Body);
        Assert.Contains("/Public", cat); Assert.Contains("/Corp", cat); Assert.Contains("/Mine", cat); Assert.DoesNotContain("/Finance", cat);
        var anon = Encoding.UTF8.GetString(srv.Handle(Req(null, null)).Body);
        Assert.Contains("/Public", anon); Assert.DoesNotContain("/Corp", anon); Assert.DoesNotContain("/Mine", anon);
    }

    [Fact]
    public void Denied_report_never_runs_the_data_function()
    {
        var store = new MemoryTemplateStore();
        var b = new TemplateBuilder("T"); b.Sheet("S", s => s.Cell("A1", "x")); store.Save(b.Build());
        var opt = new KscExcelOptions(); int calls = 0;
        opt.AddReport("T", p => { calls++; return null; }, roles: new[] { "Admin" });
        var srv = new ReportServer(opt, new ExcelReportService(opt, store));
        srv.Handle(Req("T", User("u", "Sales"), "format", "csv")); srv.Handle(Req("T", null));
        Assert.Equal(0, calls);
        srv.Handle(Req("T", User("a", "Admin"), "format", "csv"));
        Assert.Equal(1, calls);
    }

    [Fact]
    public void Row_level_security_through_the_report_context()
    {
        var (srv, _, _, _) = Setup();
        var csv = Encoding.UTF8.GetString(srv.Handle(Req("Mine", User("ali", "Sales"), "format", "csv")).Body);
        Assert.Contains("ali,10", csv); Assert.Contains("ali,5", csv); Assert.DoesNotContain("sara", csv);
        csv = Encoding.UTF8.GetString(srv.Handle(Req("Mine", User("sara", "Sales"), "format", "csv")).Body);
        Assert.Contains("sara,20", csv); Assert.DoesNotContain("ali", csv);
    }

    [Fact]
    public void Audit_records_success_denial_and_failure_without_secrets_or_messages()
    {
        var (srv, audit, opt, _) = Setup();
        opt.AddReport("Boom", p => throw new InvalidOperationException("db password is hunter2"), template: "Orders");
        srv.Handle(Req("Public", User("ali"), "format", "csv", "p_Branch", "Tehran", "p_ApiKey", "s3cr3t"));
        srv.Handle(Req("Finance", User("bob", "Sales")));
        srv.Handle(Req("Boom", null));
        srv.Handle(Req("Nope", User("ali")));                                        // unknown report: not audited (nothing to name)
        Assert.Equal(3, audit.Entries.Count);
        var ok = audit.Entries[0];
        Assert.Equal(("ali", "Public", "csv", ReportAuditStatus.Ok), (ok.User, ok.Report, ok.Action, ok.Status));
        Assert.Equal("ApiKey=***; Branch=Tehran", ok.Parameters);
        Assert.DoesNotContain("s3cr3t", ok.Parameters);
        Assert.Equal((ReportAuditStatus.Denied, "bob", "not allowed"), (audit.Entries[1].Status, audit.Entries[1].User, audit.Entries[1].Detail));
        Assert.Equal(ReportAuditStatus.Error, audit.Entries[2].Status);
        Assert.Equal("anonymous", audit.Entries[2].User);
        Assert.DoesNotContain("hunter2", audit.Entries[2].Detail);                   // the exception type, not its message
    }

    [Fact]
    public void A_failing_audit_sink_never_breaks_the_report()
    {
        var (srv, _, opt, _) = Setup();
        opt.Audit = new DelegateReportAudit(e => throw new IOException("disk full"));
        Assert.Equal(200, srv.Handle(Req("Public", null)).Status);
    }

    [Fact]
    public void File_audit_writes_one_json_line_per_event_per_day()
    {
        var dir = Path.Combine(Path.GetTempPath(), "kxl-audit-" + Guid.NewGuid().ToString("N"));
        try
        {
            var (srv, _, opt, _) = Setup();
            opt.Audit = new FileReportAudit(dir);
            srv.Handle(Req("Public", User("ali"), "format", "csv")); srv.Handle(Req("Finance", null));
            var file = Directory.GetFiles(dir).Single();
            Assert.Matches(@"audit-\d{8}\.jsonl$", file);
            var lines = File.ReadAllLines(file);
            Assert.Equal(2, lines.Length);
            var first = Newtonsoft.Json.Linq.JObject.Parse(lines[0]);
            Assert.Equal("ali", (string)first["user"]); Assert.Equal("Ok", (string)first["status"]); Assert.Equal("csv", (string)first["action"]);
            Assert.Equal("Denied", (string)Newtonsoft.Json.Linq.JObject.Parse(lines[1])["status"]);
        }
        finally { if (Directory.Exists(dir)) Directory.Delete(dir, true); }
    }

    [Fact]
    public void Scheduled_jobs_are_audited_and_bypass_role_rules()
    {
        var (_, audit, opt, svc) = Setup();
        var clock = new DateTime(2025, 3, 10, 7, 0, 0);
        var sch = new ReportScheduler(opt, svc, () => clock);
        sch.Add(new ReportJob { Name = "Nightly", Report = "Finance", Format = ReportFormat.Csv, Schedule = ReportSchedule.Daily(TimeSpan.FromHours(8)) });
        sch.RunNow("Nightly");
        var e = audit.Entries.Single();
        Assert.Equal(("scheduler", "Finance", "scheduled:csv", ReportAuditStatus.Ok), (e.User, e.Report, e.Action, e.Status));
    }
}
