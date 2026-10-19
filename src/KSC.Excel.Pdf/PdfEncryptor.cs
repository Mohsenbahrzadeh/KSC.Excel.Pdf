using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace KSC.Excel.Pdf;

/// <summary>Password protection and permissions of the generated PDF (AES-256, PDF 2.0 standard security handler, revision 6).</summary>
public sealed class PdfSecurity
{
    /// <summary>Password needed to OPEN the document. Empty = anyone can open it (only the permissions below are enforced).</summary>
    public string UserPassword { get; set; } = "";

    /// <summary>Password that lifts the restrictions (required when <see cref="UserPassword"/> is empty and any permission is denied). Random when empty.</summary>
    public string? OwnerPassword { get; set; }

    public bool AllowPrint { get; set; } = true;
    public bool AllowCopy { get; set; } = true;
    public bool AllowModify { get; set; }
    public bool AllowAnnotate { get; set; }
}

/// <summary>
/// Re-writes a PDF produced by Skia with every string and stream encrypted (AES-256-CBC, ISO 32000-2 algorithm 2.B).
/// Works on the simple structure Skia emits (classic xref table, no object streams); anything else is rejected, never half-encrypted.
/// </summary>
internal static class PdfEncryptor
{
    static readonly Encoding Latin1 = Encoding.GetEncoding("ISO-8859-1");

    public static byte[] Encrypt(byte[] pdf, PdfSecurity sec)
    {
        var rng = RandomNumberGenerator.Create();
        byte[] Rand(int n) { var b = new byte[n]; rng.GetBytes(b); return b; }

        // ---------------------------------------------------------------- keys (algorithm 8 / 9 / 10)
        var fileKey = Rand(32);
        var userPw = Pw(sec.UserPassword);
        var ownerPw = Pw(string.IsNullOrEmpty(sec.OwnerPassword) ? Convert.ToBase64String(Rand(18)) : sec.OwnerPassword!);

        var uVal = Rand(8); var uKey = Rand(8);
        var u = Concat(Hash2B(userPw, uVal, new byte[0]), uVal, uKey);                      // 48 bytes
        var ue = AesNoPad(Hash2B(userPw, uKey, new byte[0]), new byte[16], fileKey, true);   // 32 bytes
        var oVal = Rand(8); var oKey = Rand(8);
        var o = Concat(Hash2B(ownerPw, oVal, u), oVal, oKey);
        var oe = AesNoPad(Hash2B(ownerPw, oKey, u), new byte[16], fileKey, true);

        int p = unchecked((int)0xFFFFF0C0);
        if (sec.AllowPrint) p |= 4 | 2048;
        if (sec.AllowModify) p |= 8 | 256 | 1024;
        if (sec.AllowCopy) p |= 16 | 512;
        if (sec.AllowAnnotate) p |= 32 | 256;
        var perms = new byte[16];
        BitConverter.GetBytes(p).CopyTo(perms, 0);
        perms[4] = perms[5] = perms[6] = perms[7] = 0xFF; perms[8] = (byte)'T'; perms[9] = (byte)'a'; perms[10] = (byte)'d'; perms[11] = (byte)'b';
        Rand(4).CopyTo(perms, 12);
        byte[] permsEnc;
        using (var aes = Aes.Create()) { aes.Mode = CipherMode.ECB; aes.Padding = PaddingMode.None; aes.Key = fileKey; using (var enc = aes.CreateEncryptor()) permsEnc = enc.TransformFinalBlock(perms, 0, 16); }

        // ---------------------------------------------------------------- parse + rewrite
        var text = Latin1.GetString(pdf);
        int pos = text.IndexOf("obj", StringComparison.Ordinal);
        if (!text.StartsWith("%PDF-", StringComparison.Ordinal) || pos < 0) throw new InvalidOperationException("Not a PDF produced by the built-in renderer.");
        if (text.IndexOf("/ObjStm", StringComparison.Ordinal) >= 0 || text.IndexOf("/Type /XRef", StringComparison.Ordinal) >= 0 || text.IndexOf("/Type/XRef", StringComparison.Ordinal) >= 0)
            throw new NotSupportedException("PDF password protection supports classic cross-reference tables only.");

        var output = new System.IO.MemoryStream();
        void W(string s) { var b = Latin1.GetBytes(s); output.Write(b, 0, b.Length); }
        W("%PDF-1.7\n%Óëéá\n");

        var offsets = new SortedDictionary<int, long>();
        int maxObj = 0;
        int i = text.IndexOf("obj", StringComparison.Ordinal);
        // step back to the start of the first "N G obj"
        i = StartOfObject(text, i);
        while (i < text.Length)
        {
            SkipWs(text, ref i);
            if (i >= text.Length || string.CompareOrdinal(text, i, "xref", 0, 4) == 0) break;
            int num, gen; int hdrStart = i;
            if (!ReadInt(text, ref i, out num) || !ReadInt(text, ref i, out gen)) throw new InvalidOperationException("Unexpected PDF structure near offset " + i + ".");
            SkipWs(text, ref i);
            if (string.CompareOrdinal(text, i, "obj", 0, 3) != 0) throw new InvalidOperationException("Expected 'obj' at offset " + i + ".");
            i += 3;
            offsets[num] = output.Position; maxObj = Math.Max(maxObj, num);
            W(num + " " + gen + " obj");
            RewriteBody(text, pdf, ref i, output, fileKey, Rand);
            W("\nendobj\n");
        }

        // /Encrypt dictionary as one more object
        int encNum = maxObj + 1;
        offsets[encNum] = output.Position;
        W(encNum + " 0 obj\n<< /Filter /Standard /V 5 /R 6 /Length 256 /P " + p.ToString(CultureInfo.InvariantCulture)
          + " /EncryptMetadata true /CF << /StdCF << /AuthEvent /DocOpen /CFM /AESV3 /Length 32 >> >> /StmF /StdCF /StrF /StdCF"
          + " /O " + Hex(o) + " /U " + Hex(u) + " /OE " + Hex(oe) + " /UE " + Hex(ue) + " /Perms " + Hex(permsEnc) + " >>\nendobj\n");

        // original trailer: keep Root / Info
        int tr = text.LastIndexOf("trailer", StringComparison.Ordinal);
        if (tr < 0) throw new InvalidOperationException("The PDF has no trailer.");
        var trailer = text.Substring(tr);
        string Ref(string key) { var m = System.Text.RegularExpressions.Regex.Match(trailer, "/" + key + @"\s+(\d+)\s+(\d+)\s+R"); return m.Success ? "/" + key + " " + m.Groups[1].Value + " " + m.Groups[2].Value + " R " : ""; }
        var idBytes = Rand(16);

        long xrefPos = output.Position;
        W("xref\n0 " + (encNum + 1) + "\n");
        W("0000000000 65535 f \n");
        for (int n = 1; n <= encNum; n++)
        {
            long off;
            if (offsets.TryGetValue(n, out off)) W(off.ToString("0000000000", CultureInfo.InvariantCulture) + " 00000 n \n");
            else W("0000000000 65535 f \n");
        }
        W("trailer\n<< /Size " + (encNum + 1) + " " + Ref("Root") + Ref("Info") + "/Encrypt " + encNum + " 0 R /ID [" + Hex(idBytes) + " " + Hex(idBytes) + "] >>\nstartxref\n" + xrefPos + "\n%%EOF\n");
        return output.ToArray();
    }

    // ------------------------------------------------------------------------ body rewriting
    static void RewriteBody(string text, byte[] raw, ref int i, System.IO.Stream output, byte[] key, Func<int, byte[]> rand)
    {
        var sb = new StringBuilder();
        int depth = 0;
        while (i < text.Length)
        {
            char c = text[i];
            if (c == '(')
            {
                var bytes = ReadLiteral(text, ref i);
                sb.Append(Hex(EncryptData(bytes, key, rand)));
                continue;
            }
            if (c == '<' && i + 1 < text.Length && text[i + 1] == '<') { depth++; sb.Append("<<"); i += 2; continue; }
            if (c == '>' && i + 1 < text.Length && text[i + 1] == '>') { depth--; sb.Append(">>"); i += 2; continue; }
            if (c == '<')
            {
                int end = text.IndexOf('>', i);
                var hex = text.Substring(i + 1, end - i - 1);
                sb.Append(Hex(EncryptData(HexToBytes(hex), key, rand)));
                i = end + 1; continue;
            }
            if (c == '%') { while (i < text.Length && text[i] != '\n' && text[i] != '\r') i++; continue; }
            if (depth == 0 && c == 's' && string.CompareOrdinal(text, i, "stream", 0, 6) == 0 && (sb.Length == 0 || char.IsWhiteSpace(sb[sb.Length - 1]) || sb[sb.Length - 1] == '>'))
            {
                // stream: raw bytes between "stream\r?\n" and "endstream"
                int s = i + 6;
                if (s < text.Length && text[s] == '\r') s++;
                if (s < text.Length && text[s] == '\n') s++;
                int len = -1;
                var m = System.Text.RegularExpressions.Regex.Match(sb.ToString(), @"/Length\s+(\d+)(?!\s+\d+\s+R)");
                if (m.Success) len = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                int end;
                if (len >= 0 && s + len <= text.Length && string.CompareOrdinal(text, s + len, "\nendstream", 0, 10) == 0 || len >= 0 && s + len <= text.Length && string.CompareOrdinal(text, s + len, "\r\nendstream", 0, 11) == 0 || len >= 0 && s + len + 9 <= text.Length && string.CompareOrdinal(text, s + len, "endstream", 0, 9) == 0)
                    end = s + len;
                else
                {
                    end = text.IndexOf("endstream", s, StringComparison.Ordinal);
                    if (end < 0) throw new InvalidOperationException("Unterminated stream.");
                    if (end > s && text[end - 1] == '\n') end--; if (end > s && text[end - 1] == '\r') end--;
                }
                var data = new byte[end - s]; Buffer.BlockCopy(raw, s, data, 0, data.Length);
                var enc = EncryptData(data, key, rand);
                var dict = System.Text.RegularExpressions.Regex.Replace(sb.ToString(), @"/Length\s+\d+(\s+\d+\s+R)?", "/Length " + enc.Length.ToString(CultureInfo.InvariantCulture));
                if (dict.IndexOf("/Length", StringComparison.Ordinal) < 0) dict = dict.Replace("<<", "<< /Length " + enc.Length.ToString(CultureInfo.InvariantCulture) + " ");
                var head = Latin1.GetBytes(dict + "\nstream\n"); output.Write(head, 0, head.Length);
                output.Write(enc, 0, enc.Length);
                var tail = Latin1.GetBytes("\nendstream"); output.Write(tail, 0, tail.Length);
                i = text.IndexOf("endstream", end, StringComparison.Ordinal) + 9;
                sb.Clear();
                continue;
            }
            if (depth == 0 && c == 'e' && string.CompareOrdinal(text, i, "endobj", 0, 6) == 0)
            {
                var b = Latin1.GetBytes(sb.ToString()); output.Write(b, 0, b.Length);
                i += 6; return;
            }
            sb.Append(c); i++;
        }
        throw new InvalidOperationException("Unterminated object.");
    }

    static byte[] ReadLiteral(string t, ref int i)
    {
        var res = new System.IO.MemoryStream(); int depth = 0;
        i++; depth = 1;
        while (i < t.Length && depth > 0)
        {
            char c = t[i++];
            if (c == '\\')
            {
                char n = t[i++];
                switch (n)
                {
                    case 'n': res.WriteByte(10); break; case 'r': res.WriteByte(13); break; case 't': res.WriteByte(9); break;
                    case 'b': res.WriteByte(8); break; case 'f': res.WriteByte(12); break;
                    case '(': case ')': case '\\': res.WriteByte((byte)n); break;
                    case '\r': if (i < t.Length && t[i] == '\n') i++; break;
                    case '\n': break;
                    default:
                        if (n >= '0' && n <= '7')
                        {
                            int v = n - '0';
                            for (int k = 0; k < 2 && i < t.Length && t[i] >= '0' && t[i] <= '7'; k++) v = v * 8 + (t[i++] - '0');
                            res.WriteByte((byte)v);
                        }
                        else res.WriteByte((byte)n);
                        break;
                }
            }
            else if (c == '(') { depth++; res.WriteByte((byte)c); }
            else if (c == ')') { depth--; if (depth > 0) res.WriteByte((byte)c); }
            else res.WriteByte((byte)c);
        }
        return res.ToArray();
    }

    // ------------------------------------------------------------------------ crypto
    static byte[] EncryptData(byte[] data, byte[] key, Func<int, byte[]> rand)
    {
        var iv = rand(16);
        using (var aes = Aes.Create())
        {
            aes.Mode = CipherMode.CBC; aes.Padding = PaddingMode.PKCS7; aes.Key = key; aes.IV = iv;
            using (var enc = aes.CreateEncryptor())
            {
                var body = enc.TransformFinalBlock(data, 0, data.Length);
                return Concat(iv, body);
            }
        }
    }

    static byte[] AesNoPad(byte[] key, byte[] iv, byte[] data, bool encrypt)
    {
        using (var aes = Aes.Create())
        {
            aes.Mode = CipherMode.CBC; aes.Padding = PaddingMode.None; aes.Key = key; aes.IV = iv;
            using (var t = encrypt ? aes.CreateEncryptor() : aes.CreateDecryptor()) return t.TransformFinalBlock(data, 0, data.Length);
        }
    }

    /// <summary>ISO 32000-2 algorithm 2.B (revision 6 hash).</summary>
    static byte[] Hash2B(byte[] password, byte[] salt, byte[] udata)
    {
        byte[] k;
        using (var sha = SHA256.Create()) k = sha.ComputeHash(Concat(password, salt, udata));
        byte[] e = new byte[0];
        int i = 0;
        while (i < 64 || e[e.Length - 1] > i - 32)
        {
            var k1one = Concat(password, k, udata);
            var k1 = new byte[k1one.Length * 64];
            for (int r = 0; r < 64; r++) Buffer.BlockCopy(k1one, 0, k1, r * k1one.Length, k1one.Length);
            var aesKey = new byte[16]; var aesIv = new byte[16];
            Buffer.BlockCopy(k, 0, aesKey, 0, 16); Buffer.BlockCopy(k, 16, aesIv, 0, 16);
            e = AesNoPad(aesKey, aesIv, k1, true);
            int mod = 0; for (int b = 0; b < 16; b++) mod += e[b];
            mod %= 3;
            using (HashAlgorithm h = mod == 0 ? (HashAlgorithm)SHA256.Create() : mod == 1 ? SHA384.Create() : SHA512.Create()) k = h.ComputeHash(e);
            i++;
        }
        var res = new byte[32]; Buffer.BlockCopy(k, 0, res, 0, 32);
        return res;
    }

    // ------------------------------------------------------------------------ small helpers
    static byte[] Pw(string s)
    {
        var b = Encoding.UTF8.GetBytes(s ?? "");
        if (b.Length <= 127) return b;
        var t = new byte[127]; Buffer.BlockCopy(b, 0, t, 0, 127); return t;
    }

    static byte[] Concat(params byte[][] parts)
    {
        int n = 0; foreach (var p in parts) n += p.Length;
        var r = new byte[n]; int o = 0;
        foreach (var p in parts) { Buffer.BlockCopy(p, 0, r, o, p.Length); o += p.Length; }
        return r;
    }

    static string Hex(byte[] b) { var sb = new StringBuilder(b.Length * 2 + 2); sb.Append('<'); foreach (var x in b) sb.Append(x.ToString("X2", CultureInfo.InvariantCulture)); sb.Append('>'); return sb.ToString(); }

    static byte[] HexToBytes(string hex)
    {
        var clean = new StringBuilder(); foreach (var c in hex) if (!char.IsWhiteSpace(c)) clean.Append(c);
        if (clean.Length % 2 == 1) clean.Append('0');
        var r = new byte[clean.Length / 2];
        for (int i = 0; i < r.Length; i++) r[i] = byte.Parse(clean.ToString(i * 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return r;
    }

    static void SkipWs(string t, ref int i) { while (i < t.Length && (char.IsWhiteSpace(t[i]) || t[i] == '\0')) i++; }

    static bool ReadInt(string t, ref int i, out int v)
    {
        SkipWs(t, ref i); int s = i; v = 0;
        while (i < t.Length && char.IsDigit(t[i])) i++;
        return i > s && int.TryParse(t.Substring(s, i - s), NumberStyles.None, CultureInfo.InvariantCulture, out v);
    }

    static int StartOfObject(string t, int objKeywordPos)
    {
        int i = objKeywordPos - 1;
        while (i >= 0 && char.IsWhiteSpace(t[i])) i--;
        while (i >= 0 && char.IsDigit(t[i])) i--;     // generation
        while (i >= 0 && char.IsWhiteSpace(t[i])) i--;
        while (i >= 0 && char.IsDigit(t[i])) i--;     // number
        return i + 1;
    }
}
