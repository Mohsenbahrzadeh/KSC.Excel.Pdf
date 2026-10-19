using System.Collections.ObjectModel;
using System.Data;
using KSC.Excel.Designer;
using KSC.Excel.Engine;
using KSC.Excel.Model;
using KSC.Excel.Services;
using Xunit;

namespace KSC.Excel.Tests;

// ---- a tiny fake ADO.NET provider backed by a DataTable (no database needed)
sealed class FakeParams : Collection<IDbDataParameter>, IDataParameterCollection
{
    public object this[string name]
    {
        get => this.First(p => p.ParameterName == name);
        set { }
    }
    public bool Contains(string name) => this.Any(p => p.ParameterName == name);
    public int IndexOf(string name) => IndexOf(this.FirstOrDefault(p => p.ParameterName == name));
    public void RemoveAt(string name) => Remove(this.First(p => p.ParameterName == name));
}

sealed class FakeParam : IDbDataParameter
{
    public DbType DbType { get; set; }
    public ParameterDirection Direction { get; set; }
    public bool IsNullable => true;
    public string ParameterName { get; set; }
    public string SourceColumn { get; set; }
    public DataRowVersion SourceVersion { get; set; }
    public object Value { get; set; }
    public byte Precision { get; set; }
    public byte Scale { get; set; }
    public int Size { get; set; }
}

sealed class FakeCommand : IDbCommand
{
    readonly DataTable _table; readonly List<string> _log;
    public FakeCommand(DataTable t, List<string> log) { _table = t; _log = log; }
    public string CommandText { get; set; }
    public int CommandTimeout { get; set; }
    public CommandType CommandType { get; set; }
    public IDbConnection Connection { get; set; }
    public IDataParameterCollection Parameters { get; } = new FakeParams();
    public IDbTransaction Transaction { get; set; }
    public UpdateRowSource UpdatedRowSource { get; set; }
    public void Cancel() { }
    public IDbDataParameter CreateParameter() => new FakeParam();
    public void Dispose() { }
    public int ExecuteNonQuery() => throw new InvalidOperationException("write attempted");
    public IDataReader ExecuteReader() { _log.Add(CommandText + " | " + string.Join(",", Parameters.Cast<IDbDataParameter>().Select(p => p.ParameterName + "=" + p.Value))); return _table.CreateDataReader(); }
    public IDataReader ExecuteReader(CommandBehavior behavior) => ExecuteReader();
    public object ExecuteScalar() => null;
    public void Prepare() { }
}

sealed class FakeConnection : IDbConnection
{
    readonly DataTable _table; readonly List<string> _log;
    public FakeConnection(DataTable t, List<string> log) { _table = t; _log = log; }
    public string ConnectionString { get; set; }
    public int ConnectionTimeout => 15;
    public string Database => "fake";
    public ConnectionState State { get; private set; } = ConnectionState.Closed;
    public IDbTransaction BeginTransaction() => throw new NotSupportedException();
    public IDbTransaction BeginTransaction(IsolationLevel il) => throw new NotSupportedException();
    public void ChangeDatabase(string databaseName) { }
    public void Close() { State = ConnectionState.Closed; }
    public IDbCommand CreateCommand() => new FakeCommand(_table, _log);
    public void Open() { State = ConnectionState.Open; }
    public void Dispose() { Close(); }
}

public class SqlAndExprTests
{
    static DataTable Orders()
    {
        var t = new DataTable();
        t.Columns.Add("Customer"); t.Columns.Add("Amount", typeof(decimal)); t.Columns.Add("Day", typeof(DateTime));
        t.Rows.Add("Ali", 100m, new DateTime(2025, 3, 21)); t.Rows.Add("Sara", 250m, new DateTime(2025, 3, 22)); t.Rows.Add("Ali", 50m, new DateTime(2025, 3, 23));
        return t;
    }

    static (IExcelReportService svc, List<string> log) Service(bool allow, Action<KscExcelOptions> more = null)
    {
        var log = new List<string>();
        var o = new KscExcelOptions { AllowSql = allow };
        o.AddConnection("Main", () => new FakeConnection(Orders(), log));
        more?.Invoke(o);
        return (new ExcelReportService(o, new MemoryTemplateStore()), log);
    }

    static ExcelTemplate SqlTemplate(string sql = "SELECT Customer, Amount, Day FROM Orders WHERE Amount >= @Min")
    {
        var b = new TemplateBuilder("Q");
        b.Sheet("S", s => s.Cell("A1", "{{Customer}}").Cell("B1", "{{Amount}}").Cell("C1", "{{Amount * 2}}"));
        var t = b.Build();
        t.Parameters.Add(new ParameterDef { Name = "Min", Type = "number", Default = "60" });
        t.DataSets.Add(new DataSetDef { Name = "Orders", Type = "sql", Connection = "Main", Sql = sql });
        t.Sheets[0].Bands.Add(new BandDef { Name = "O", Source = "Orders", FirstRow = 1, LastRow = 1 });
        return t;
    }

    [Fact]
    public void Sql_data_set_binds_parameters_and_feeds_a_band()
    {
        var (svc, log) = Service(true);
        var x = Xlsx.Open(KSC.Excel.Writer.XlsxWriter.ToBytes(svc.Bind(SqlTemplate(), null, new Dictionary<string, object> { ["min"] = 80 })));
        Assert.Equal("Ali", x.Value("A1")); Assert.Equal("100", x.Value("B1")); Assert.Equal("200", x.Value("C1"));
        Assert.Equal("Sara", x.Value("A2"));
        Assert.Contains("@Min=80", log.Single());                       // bound as a parameter, not concatenated
        Assert.DoesNotContain("80 ", log.Single().Split('|')[0]);
    }

    [Fact]
    public void Sql_disabled_throws_at_runtime_but_preview_skips()
    {
        var (svc, _) = Service(false);
        Assert.Throws<ExcelRenderException>(() => svc.Bind(SqlTemplate(), null, null));
        var wb = svc.BindPreview(SqlTemplate(), Make.Json("{\"Orders\":[{\"Customer\":\"Sample\",\"Amount\":1}]}"), null);
        Assert.Contains(wb.Warnings, w => w.Contains("not executed"));
        Assert.Equal("Sample", wb.Sheets[0].Get(1, 1).Value);          // falls back to sample data
    }

    [Theory]
    [InlineData("DELETE FROM Orders")]
    [InlineData("SELECT 1; DROP TABLE Orders")]
    [InlineData("SELECT * FROM Orders; SELECT 2")]
    [InlineData("EXEC sp_who")]
    [InlineData("SELECT * INTO #x FROM Orders")]
    [InlineData("UPDATE Orders SET Amount = 0")]
    [InlineData("")]
    public void Sql_guard_rejects_non_select_statements(string sql)
        => Assert.Throws<ExcelRenderException>(() => SqlDataSource.Validate(sql));

    [Theory]
    [InlineData("SELECT * FROM Orders")]
    [InlineData("  with x as (select 1 as a) select a from x;  ")]
    [InlineData("SELECT 'drop table; delete' AS note -- insert here\n FROM Orders")]
    public void Sql_guard_allows_selects_even_with_keywords_in_literals_and_comments(string sql)
        => SqlDataSource.Validate(sql);

    [Fact]
    public void Sql_max_rows_is_enforced()
    {
        var (svc, _) = Service(true, o => o.SqlMaxRows = 2);
        Assert.Throws<ExcelRenderException>(() => svc.Bind(SqlTemplate(), null, null));
    }

    [Fact]
    public void Pivot_data_set_can_use_a_sql_data_set()
    {
        var (svc, _) = Service(true);
        var b = new TemplateBuilder("P");
        b.Sheet("S", s => s.Cell("A1", "{{GrandTotal}}").Cell("A2", "{{Matrix.GrandTotal}}"));
        var t = b.Build();
        t.DataSets.Add(new DataSetDef { Name = "Orders", Type = "sql", Connection = "Main", Sql = "SELECT * FROM Orders" });
        t.DataSets.Add(new DataSetDef { Name = "Matrix", Type = "pivot", Source = "Orders", RowPath = "Customer", ColPath = "Day", ValuePath = "Amount" });
        var x = Xlsx.Open(KSC.Excel.Writer.XlsxWriter.ToBytes(svc.Bind(t, null, null)));
        Assert.Equal("400", x.Value("A2"));
    }

    // ---------------------------------------------------------------- expressions & custom functions
    [Fact]
    public void Expressions_calculated_fields()
    {
        var b = new TemplateBuilder("E");
        b.Sheet("S", s => s
            .Cell("A1", "{{Qty * Price}}").Cell("A2", "{{IF(Qty > 5, 'bulk', 'single')}}").Cell("A3", "{{ROUND(Price * 1.09, 1)}}")
            .Cell("A4", "{{Name & '-' & Qty}}").Cell("A5", "{{(Qty + 1) * 2 - 3}}").Cell("A6", "{{COALESCE(Missing, 'n/a')}}")
            .Cell("A7", "{{Sum(Items.Qty) + 1}}").Cell("A8", "{{Price / 0}}").Cell("A9", "Total: {{Qty * 2}}"));
        var x = Make.Render(b.Build(), Make.Json("{\"Qty\":10,\"Price\":3.5,\"Name\":\"ab\",\"Items\":[{\"Qty\":2},{\"Qty\":3}]}"));
        Assert.Equal("35", x.Value("A1")); Assert.Equal("bulk", x.Value("A2")); Assert.Equal("3.8", x.Value("A3"));
        Assert.Equal("ab-10", x.Value("A4")); Assert.Equal("19", x.Value("A5")); Assert.Equal("n/a", x.Value("A6"));
        Assert.Equal("6", x.Value("A7")); Assert.Null(x.Value("A8")); Assert.Equal("Total: 20", x.Value("A9"));
    }

    [Fact]
    public void Custom_functions_and_filters_are_available_in_placeholders_and_formulas()
    {
        var o = new KscExcelOptions();
        o.AddFunction("TAX", a => Convert.ToDouble(a[0]) * 0.09);
        o.AddFunction("SUMSQ2", a => ((object[])a[0]).Sum(v => Convert.ToDouble(v) * Convert.ToDouble(v)));
        o.AddFilter("stars", (v, arg) => new string('*', Convert.ToInt32(v)));
        var svc = new ExcelReportService(o, new MemoryTemplateStore());
        var b = new TemplateBuilder("C");
        b.Sheet("S", s => s.Cell("A1", "{{Tax(Price)}}").Cell("A2", "{{Price + Tax(Price)}}").Cell("A3", "{{N|stars}}")
            .Cell("B1", "1").Cell("B2", "2").Cell("B3", "3").Cell("C1", "=SUMSQ2(B1:B3)").Cell("C2", "=TAX(B3)+1"));
        var x = Xlsx.Open(KSC.Excel.Writer.XlsxWriter.ToBytes(svc.Bind(b.Build(), Make.Json("{\"Price\":100,\"N\":3}"), null)));
        Assert.Equal("9", x.Value("A1")); Assert.Equal("109", x.Value("A2")); Assert.Equal("***", x.Value("A3"));
        Assert.Equal("14", x.Value("C1"));                 // cached values computed through the custom Excel function
        Assert.Equal("1.27", x.Value("C2"));
        Assert.Equal("SUMSQ2(B1:B3)", x.Formula("C1"));    // formula text is untouched for Excel
    }

    [Fact]
    public void Unknown_function_in_expression_warns_instead_of_crashing()
    {
        var b = new TemplateBuilder("W");
        b.Sheet("S", s => s.Cell("A1", "{{Nope(1)}}"));
        var wb = new TemplateRenderer().Render(b.Build(), null);
        Assert.Contains(wb.Warnings, w => w.Contains("NOPE"));
    }
}

public class SqlDesignerTests
{
    static DesignerRequest Post(string path, string body)
    {
        var r = new DesignerRequest { Method = "POST", Path = path, BasePath = "/d", Body = System.Text.Encoding.UTF8.GetBytes(body) };
        r.Headers["X-KXL-Request"] = "1";
        return r;
    }

    [Fact]
    public void Sql_test_endpoint_is_forbidden_unless_enabled_and_returns_columns_when_enabled()
    {
        var off = new KscExcelOptions();
        var hostOff = new DesignerHost(off, new ExcelReportService(off, new MemoryTemplateStore()));
        Assert.Equal(403, hostOff.Handle(Post("/api/sql/test", "{\"connection\":\"Main\",\"sql\":\"SELECT 1\"}")).Status);

        var table = new DataTable(); table.Columns.Add("Name"); table.Columns.Add("Qty", typeof(int)); table.Rows.Add("a", 1);
        var on = new KscExcelOptions { AllowSql = true };
        on.AddConnection("Main", () => new FakeConnection(table, new List<string>()));
        var host = new DesignerHost(on, new ExcelReportService(on, new MemoryTemplateStore()));
        var res = host.Handle(Post("/api/sql/test", "{\"connection\":\"Main\",\"sql\":\"SELECT Name, Qty FROM T\"}"));
        Assert.Equal(200, res.Status);
        var json = Newtonsoft.Json.Linq.JObject.Parse(System.Text.Encoding.UTF8.GetString(res.Body));
        Assert.Equal("Qty", (string)json["columns"][1]["name"]); Assert.Equal("number", (string)json["columns"][1]["type"]);
        Assert.Equal(400, host.Handle(Post("/api/sql/test", "{\"connection\":\"Nope\",\"sql\":\"SELECT 1\"}")).Status);
        Assert.Equal(422, host.Handle(Post("/api/sql/test", "{\"connection\":\"Main\",\"sql\":\"DROP TABLE x\"}")).Status);   // guard
    }
}
