# Security policy

## Supported versions

The latest released minor version receives security fixes.

## Reporting a vulnerability

Please **do not open a public issue** for security problems.
Use GitHub's private vulnerability reporting ("Security" tab → "Report a vulnerability") or contact the maintainers privately.
We aim to acknowledge reports within 5 working days.

## Things worth knowing when you deploy

- **The designer can write templates.** Always put it behind authentication (`RequireAuthorization(...)` / `KscExcel.Authorize`) and consider `ReadOnlyDesigner = true` in production.
- **SQL data sets are off by default** (`AllowSql`). When enabled, use a **read-only** database login: the engine rejects anything except a single `SELECT`/`WITH`, but that is a second line of defence.
- **REST data sets** can only reach hosts registered in code with `AddHttpEndpoint` (templates choose a path, never a host). **Providers** (`AddProvider`) run your own code.
- User data is never turned into formulas (no formula/CSV injection); the HTML preview is fully encoded.
- Write requests to the designer API require the `X-KXL-Request: 1` header (CSRF guard).
