using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace KSC.Excel.Engine
{
    /// <summary>A barcode as a module matrix (true = dark). 1-D codes have a single row.</summary>
    public sealed class BarcodeMatrix
    {
        public bool[,] Modules;          // [row, col]
        public int Rows { get { return Modules.GetLength(0); } }
        public int Cols { get { return Modules.GetLength(1); } }
        /// <summary>Quiet-zone width in modules that should surround the code.</summary>
        public int Quiet;
        public bool Is2D;
    }

    /// <summary>
    /// Dependency-free barcode generation: QR Code (byte mode, versions 1-10, ECC L/M/Q/H) and Code 128 (sets B and C), plus a PNG writer
    /// for the .xlsx output. The PDF renderer draws the same matrix as vector rectangles.
    /// </summary>
    public static class Barcodes
    {
        public static readonly string[] Types = { "qr", "code128", "ean13", "ean8", "code39" };

        public static bool IsSupported(string type)
        {
            if (string.IsNullOrEmpty(type)) return false;
            var t = type.ToLowerInvariant();
            return t == "qr" || t.StartsWith("qr:") || t == "code128" || t == "ean13" || t == "ean8" || t == "code39";
        }

        /// <summary>Builds the module matrix. "qr" (optionally "qr:L|M|Q|H") or "code128". Throws ArgumentException when the text cannot be encoded.</summary>
        public static BarcodeMatrix Encode(string type, string text)
        {
            if (string.IsNullOrEmpty(text)) throw new ArgumentException("The barcode text is empty.");
            var t = (type ?? "").ToLowerInvariant();
            if (t == "qr" || t.StartsWith("qr:"))
            {
                var ecl = t.Length > 3 ? char.ToUpperInvariant(t[3]) : 'M';
                var q = QrEncoder.Encode(Encoding.UTF8.GetBytes(text), ecl);
                return new BarcodeMatrix { Modules = q, Quiet = 2, Is2D = true };
            }
            if (t == "code128")
            {
                var bars = Code128.Encode(text);
                var m = new bool[1, bars.Length];
                for (int i = 0; i < bars.Length; i++) m[0, i] = bars[i];
                return new BarcodeMatrix { Modules = m, Quiet = 10, Is2D = false };
            }
            if (t == "ean13" || t == "ean8" || t == "code39")
            {
                var bars = t == "ean13" ? Linear.Ean13(text) : t == "ean8" ? Linear.Ean8(text) : Linear.Code39(text);
                var m = new bool[1, bars.Length];
                for (int i = 0; i < bars.Length; i++) m[0, i] = bars[i];
                return new BarcodeMatrix { Modules = m, Quiet = t == "code39" ? 10 : 9, Is2D = false };
            }
            throw new ArgumentException("Unknown barcode type '" + type + "'. Use qr or code128.");
        }

        /// <summary>Renders a matrix to a PNG of about <paramref name="maxW"/> x <paramref name="maxH"/> pixels (never larger); returns the actual size.</summary>
        public static byte[] ToPng(BarcodeMatrix m, int maxW, int maxH, out int width, out int height)
        {
            int cols = m.Cols + 2 * m.Quiet, rows = m.Is2D ? m.Rows + 2 * m.Quiet : 1;
            int k = Math.Max(1, Math.Min(maxW / cols, m.Is2D ? maxH / rows : maxW));
            width = cols * k;
            height = m.Is2D ? rows * k : Math.Max(8, maxH);
            var pixels = new byte[width * height];
            for (int i = 0; i < pixels.Length; i++) pixels[i] = 255;
            for (int r = 0; r < (m.Is2D ? m.Rows : 1); r++)
                for (int c = 0; c < m.Cols; c++)
                {
                    if (!m.Modules[r, c]) continue;
                    int x0 = (c + m.Quiet) * k;
                    int y0 = m.Is2D ? (r + m.Quiet) * k : 0, y1 = m.Is2D ? y0 + k : height;
                    for (int y = y0; y < y1; y++)
                        for (int x = x0; x < x0 + k; x++) pixels[y * width + x] = 0;
                }
            return Png.Gray8(pixels, width, height);
        }

        // =================================================================== EAN-13 / EAN-8 / Code 39
        static class Linear
        {
            static readonly string[] L = { "0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011" };
            static readonly string[] G = { "0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111" };
            static readonly string[] R = { "1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100" };
            static readonly string[] Parity = { "LLLLLL", "LLGLGG", "LLGGLG", "LLGGGL", "LGLLGG", "LGGLLG", "LGGGLL", "LGLGLG", "LGLGGL", "LGGLGL" };

            static int CheckDigit(string digits)
            {
                int sum = 0;
                for (int i = 0; i < digits.Length; i++)
                {
                    int d = digits[digits.Length - 1 - i] - '0';
                    sum += (i % 2 == 0) ? d * 3 : d;
                }
                return (10 - sum % 10) % 10;
            }

            static string Digits(string text, int bodyLen, string name)
            {
                var t = (text ?? "").Trim();
                foreach (var ch in t) if (ch < '0' || ch > '9') throw new ArgumentException(name + " needs digits only.");
                if (t.Length == bodyLen) t += CheckDigit(t);
                else if (t.Length == bodyLen + 1)
                {
                    if (t[bodyLen] - '0' != CheckDigit(t.Substring(0, bodyLen))) throw new ArgumentException(name + ": wrong check digit (expected " + CheckDigit(t.Substring(0, bodyLen)) + ").");
                }
                else throw new ArgumentException(name + " needs " + bodyLen + " digits (the check digit is added) or " + (bodyLen + 1) + " with a valid check digit.");
                return t;
            }

            static bool[] Bits(string s) { var b = new bool[s.Length]; for (int i = 0; i < s.Length; i++) b[i] = s[i] == '1'; return b; }

            public static bool[] Ean13(string text)
            {
                var d = Digits(text, 12, "EAN-13");
                var sb = new StringBuilder("101");
                var par = Parity[d[0] - '0'];
                for (int i = 1; i <= 6; i++) sb.Append(par[i - 1] == 'L' ? L[d[i] - '0'] : G[d[i] - '0']);
                sb.Append("01010");
                for (int i = 7; i <= 12; i++) sb.Append(R[d[i] - '0']);
                sb.Append("101");
                return Bits(sb.ToString());
            }

            public static bool[] Ean8(string text)
            {
                var d = Digits(text, 7, "EAN-8");
                var sb = new StringBuilder("101");
                for (int i = 0; i < 4; i++) sb.Append(L[d[i] - '0']);
                sb.Append("01010");
                for (int i = 4; i < 8; i++) sb.Append(R[d[i] - '0']);
                sb.Append("101");
                return Bits(sb.ToString());
            }

            // bar/space widths (n = narrow, w = wide), 9 elements per character
            const string Alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*";
            static readonly string[] C39 =
            {
                "nnnwwnwnn", "wnnwnnnnw", "nnwwnnnnw", "wnwwnnnnn", "nnnwwnnnw", "wnnwwnnnn", "nnwwwnnnn", "nnnwnnwnw", "wnnwnnwnn", "nnwwnnwnn",
                "wnnnnwnnw", "nnwnnwnnw", "wnwnnwnnn", "nnnnwwnnw", "wnnnwwnnn", "nnwnwwnnn", "nnnnnwwnw", "wnnnnwwnn", "nnwnnwwnn", "nnnnwwwnn",
                "wnnnnnnww", "nnwnnnnww", "wnwnnnnwn", "nnnnwnnww", "wnnnwnnwn", "nnwnwnnwn", "nnnnnnwww", "wnnnnnwwn", "nnwnnnwwn", "nnnnwnwwn",
                "wwnnnnnnw", "nwwnnnnnw", "wwwnnnnnn", "nwnnwnnnw", "wwnnwnnnn", "nwwnwnnnn",
                "nwnnnnwnw", "wwnnnnwnn", "nwwnnnwnn", "nwnwnwnnn", "nwnwnnnwn", "nwnnnwnwn", "nnnwnwnwn", "nwnnwnwnn"
            };

            public static bool[] Code39(string text)
            {
                var t = (text ?? "").ToUpperInvariant();
                foreach (var ch in t) if (Alphabet.IndexOf(ch) < 0 || ch == '*') throw new ArgumentException("Code 39 supports 0-9 A-Z - . space $ / + % (got '" + ch + "').");
                var bits = new List<bool>();
                foreach (var ch in "*" + t + "*")
                {
                    var pat = C39[Alphabet.IndexOf(ch)];
                    for (int i = 0; i < 9; i++)
                    {
                        bool bar = i % 2 == 0; int w = pat[i] == 'w' ? 3 : 1;
                        for (int k = 0; k < w; k++) bits.Add(bar);
                    }
                    bits.Add(false);   // inter-character gap
                }
                bits.RemoveAt(bits.Count - 1);
                return bits.ToArray();
            }
        }

        // =================================================================== Code 128
        static class Code128
        {
            static readonly string[] Patterns =
            {
                "212222","222122","222221","121223","121322","131222","122213","122312","132212","221213",
                "221312","231212","112232","122132","122231","113222","123122","123221","223211","221132",
                "221231","213212","223112","312131","311222","321122","321221","312212","322112","322211",
                "212123","212321","232121","111323","131123","131321","112313","132113","132311","211313",
                "231113","231311","112133","112331","132131","113123","113321","133121","313121","211331",
                "231131","213113","213311","213131","311123","311321","331121","312113","312311","332111",
                "314111","221411","431111","111224","111422","121124","121421","141122","141221","112214",
                "112412","122114","122411","142112","142211","241211","221114","413111","241112","134111",
                "111242","121142","121241","114212","124112","124211","411212","421112","421211","212141",
                "214121","412121","111143","111341","131141","114113","114311","411113","411311","113141",
                "114131","311141","411131","211412","211214","211232","2331112"
            };
            internal static string[] AllPatterns { get { return Patterns; } }

            public static bool[] Encode(string text)
            {
                var values = new List<int>();
                bool allDigits = text.Length >= 4 && text.Length % 2 == 0;
                foreach (var ch in text) { if (ch < 32 || ch > 126) throw new ArgumentException("Code 128 can only encode printable ASCII characters (got U+" + ((int)ch).ToString("X4") + ")."); if (ch < '0' || ch > '9') allDigits = false; }
                if (allDigits)
                {
                    values.Add(105);
                    for (int i = 0; i < text.Length; i += 2) values.Add((text[i] - '0') * 10 + (text[i + 1] - '0'));
                }
                else
                {
                    values.Add(104);
                    foreach (var ch in text) values.Add(ch - 32);
                }
                long sum = values[0];
                for (int i = 1; i < values.Count; i++) sum += (long)values[i] * i;
                values.Add((int)(sum % 103));
                values.Add(106);

                var bars = new List<bool>();
                foreach (var v in values)
                {
                    bool dark = true;
                    foreach (var w in Patterns[v]) { for (int i = 0; i < w - '0'; i++) bars.Add(dark); dark = !dark; }
                }
                return bars.ToArray();
            }
        }

        /// <summary>Exposed for tests: the bar/space width patterns of Code 128 (each 11 modules, the stop 13).</summary>
        public static string[] Code128Patterns { get { return Code128.AllPatterns; } }

        // =================================================================== QR Code
        static class QrEncoder
        {
            static readonly int[][] EccPerBlock =
            {
                new[] { 7, 10, 15, 20, 26, 18, 20, 24, 30, 18 },   // L
                new[] { 10, 16, 26, 18, 24, 16, 18, 22, 22, 26 },  // M
                new[] { 13, 22, 18, 26, 18, 24, 18, 22, 20, 24 },  // Q
                new[] { 17, 28, 22, 16, 22, 28, 26, 26, 24, 28 }   // H
            };
            static readonly int[][] NumBlocks =
            {
                new[] { 1, 1, 1, 1, 1, 2, 2, 2, 2, 4 },
                new[] { 1, 1, 1, 2, 2, 4, 4, 4, 5, 5 },
                new[] { 1, 1, 2, 2, 4, 4, 6, 6, 8, 8 },
                new[] { 1, 1, 2, 4, 4, 4, 5, 6, 8, 8 }
            };
            static readonly int[] FormatBits = { 1, 0, 3, 2 };   // L M Q H
            const int MaxVersion = 10;

            static int RawModules(int ver)
            {
                int r = (16 * ver + 128) * ver + 64;
                if (ver >= 2)
                {
                    int na = ver / 7 + 2;
                    r -= (25 * na - 10) * na - 55;
                    if (ver >= 7) r -= 36;
                }
                return r;
            }

            public static bool[,] Encode(byte[] data, char eccLevel)
            {
                int e = "LMQH".IndexOf(eccLevel);
                if (e < 0) e = 1;
                int ver = 0, dataCodewords = 0;
                for (int v = 1; v <= MaxVersion; v++)
                {
                    int cap = RawModules(v) / 8 - EccPerBlock[e][v - 1] * NumBlocks[e][v - 1];
                    int need = (4 + (v < 10 ? 8 : 16) + 8 * data.Length + 7) / 8;
                    if (need <= cap) { ver = v; dataCodewords = cap; break; }
                }
                if (ver == 0) throw new ArgumentException("The QR text is too long (" + data.Length + " bytes). The built-in encoder supports up to about 270 bytes; shorten it or use a lower ECC level (qr:L).");

                // bit stream
                var bits = new List<bool>();
                Append(bits, 4, 4);                                  // byte mode
                Append(bits, data.Length, ver < 10 ? 8 : 16);
                foreach (var b in data) Append(bits, b, 8);
                int capBits = dataCodewords * 8;
                for (int i = 0; i < 4 && bits.Count < capBits; i++) bits.Add(false);
                while (bits.Count % 8 != 0) bits.Add(false);
                for (int pad = 0xEC; bits.Count < capBits; pad ^= 0xEC ^ 0x11) Append(bits, pad, 8);
                var cw = new byte[dataCodewords];
                for (int i = 0; i < bits.Count; i++) if (bits[i]) cw[i >> 3] |= (byte)(1 << (7 - (i & 7)));

                var all = AddEccAndInterleave(cw, ver, e);
                return Build(all, ver, e);
            }

            static void Append(List<bool> bits, int val, int len) { for (int i = len - 1; i >= 0; i--) bits.Add(((val >> i) & 1) != 0); }

            static byte[] AddEccAndInterleave(byte[] data, int ver, int e)
            {
                int numBlocks = NumBlocks[e][ver - 1], blockEcc = EccPerBlock[e][ver - 1];
                int raw = RawModules(ver) / 8;
                int shortBlocks = numBlocks - raw % numBlocks, shortLen = raw / numBlocks;
                var blocks = new byte[numBlocks][];
                var rsDiv = RsDivisor(blockEcc);
                for (int i = 0, k = 0; i < numBlocks; i++)
                {
                    int datLen = shortLen - blockEcc + (i < shortBlocks ? 0 : 1);
                    var dat = new byte[datLen];
                    Array.Copy(data, k, dat, 0, datLen); k += datLen;
                    var ecc = RsRemainder(dat, rsDiv);
                    var block = new byte[shortLen + 1];
                    Array.Copy(dat, block, datLen);
                    Array.Copy(ecc, 0, block, shortLen + 1 - blockEcc, blockEcc);
                    blocks[i] = block;
                }
                var result = new byte[raw];
                for (int i = 0, k = 0; i < blocks[0].Length; i++)
                    for (int j = 0; j < blocks.Length; j++)
                        if (i != shortLen - blockEcc || j >= shortBlocks) result[k++] = blocks[j][i];
                return result;
            }

            static byte[] RsDivisor(int degree)
            {
                var result = new byte[degree];
                result[degree - 1] = 1;
                int root = 1;
                for (int i = 0; i < degree; i++)
                {
                    for (int j = 0; j < degree; j++)
                    {
                        result[j] = (byte)Mul(result[j], root);
                        if (j + 1 < degree) result[j] ^= result[j + 1];
                    }
                    root = Mul(root, 0x02);
                }
                return result;
            }

            static byte[] RsRemainder(byte[] data, byte[] divisor)
            {
                var result = new byte[divisor.Length];
                foreach (var b in data)
                {
                    int factor = b ^ result[0];
                    Array.Copy(result, 1, result, 0, result.Length - 1);
                    result[result.Length - 1] = 0;
                    for (int i = 0; i < result.Length; i++) result[i] ^= (byte)Mul(divisor[i], factor);
                }
                return result;
            }

            static int Mul(int x, int y)
            {
                int z = 0;
                for (int i = 7; i >= 0; i--)
                {
                    z = (z << 1) ^ ((z >> 7) * 0x11D);
                    z ^= ((y >> i) & 1) * x;
                }
                return z;
            }

            static bool Bit(int x, int i) { return ((x >> i) & 1) != 0; }

            sealed class Grid
            {
                public int Size; public bool[,] M; public bool[,] Fn;
                public Grid(int size) { Size = size; M = new bool[size, size]; Fn = new bool[size, size]; }
                public void Set(int x, int y, bool dark) { M[y, x] = dark; Fn[y, x] = true; }
            }

            static bool[,] Build(byte[] codewords, int ver, int e)
            {
                int size = ver * 4 + 17;
                var g = new Grid(size);
                for (int i = 0; i < size; i++) { g.Set(6, i, i % 2 == 0); g.Set(i, 6, i % 2 == 0); }
                Finder(g, 3, 3); Finder(g, size - 4, 3); Finder(g, 3, size - 4);
                var pos = AlignPositions(ver);
                for (int i = 0; i < pos.Length; i++)
                    for (int j = 0; j < pos.Length; j++)
                        if (!((i == 0 && j == 0) || (i == 0 && j == pos.Length - 1) || (i == pos.Length - 1 && j == 0))) Align(g, pos[i], pos[j]);
                DrawFormat(g, e, 0);
                DrawVersion(g, ver);

                // data
                int idx = 0;
                for (int right = size - 1; right >= 1; right -= 2)
                {
                    if (right == 6) right = 5;
                    for (int vert = 0; vert < size; vert++)
                        for (int j = 0; j < 2; j++)
                        {
                            int x = right - j;
                            bool upward = ((right + 1) & 2) == 0;
                            int y = upward ? size - 1 - vert : vert;
                            if (!g.Fn[y, x] && idx < codewords.Length * 8) { g.M[y, x] = Bit(codewords[idx >> 3], 7 - (idx & 7)); idx++; }
                        }
                }

                int best = 0; long bestPenalty = long.MaxValue;
                for (int mask = 0; mask < 8; mask++)
                {
                    ApplyMask(g, mask); DrawFormat(g, e, mask);
                    long p = Penalty(g);
                    if (p < bestPenalty) { bestPenalty = p; best = mask; }
                    ApplyMask(g, mask);   // undo
                }
                ApplyMask(g, best); DrawFormat(g, e, best);
                return g.M;
            }

            static void Finder(Grid g, int cx, int cy)
            {
                for (int dy = -4; dy <= 4; dy++)
                    for (int dx = -4; dx <= 4; dx++)
                    {
                        int d = Math.Max(Math.Abs(dx), Math.Abs(dy)), x = cx + dx, y = cy + dy;
                        if (x >= 0 && x < g.Size && y >= 0 && y < g.Size) g.Set(x, y, d != 2 && d != 4);
                    }
            }

            static void Align(Grid g, int cx, int cy)
            {
                for (int dy = -2; dy <= 2; dy++)
                    for (int dx = -2; dx <= 2; dx++) g.Set(cx + dx, cy + dy, Math.Max(Math.Abs(dx), Math.Abs(dy)) != 1);
            }

            static int[] AlignPositions(int ver)
            {
                if (ver == 1) return new int[0];
                int n = ver / 7 + 2;
                int step = ver == 32 ? 26 : (ver * 8 + n * 3 + 5) / (n * 4 - 4) * 2;
                var r = new int[n];
                r[0] = 6;
                for (int i = n - 1, pos = ver * 4 + 10; i >= 1; i--, pos -= step) r[i] = pos;
                return r;
            }

            static void DrawFormat(Grid g, int e, int mask)
            {
                int data = FormatBits[e] << 3 | mask, rem = data;
                for (int i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >> 9) * 0x537);
                int bits = (data << 10 | rem) ^ 0x5412;
                int s = g.Size;
                for (int i = 0; i <= 5; i++) g.Set(8, i, Bit(bits, i));
                g.Set(8, 7, Bit(bits, 6)); g.Set(8, 8, Bit(bits, 7)); g.Set(7, 8, Bit(bits, 8));
                for (int i = 9; i < 15; i++) g.Set(14 - i, 8, Bit(bits, i));
                for (int i = 0; i < 8; i++) g.Set(s - 1 - i, 8, Bit(bits, i));
                for (int i = 8; i < 15; i++) g.Set(8, s - 15 + i, Bit(bits, i));
                g.Set(8, s - 8, true);
            }

            static void DrawVersion(Grid g, int ver)
            {
                if (ver < 7) return;
                int rem = ver;
                for (int i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >> 11) * 0x1F25);
                int bits = ver << 12 | rem;
                for (int i = 0; i < 18; i++)
                {
                    bool bit = Bit(bits, i);
                    int a = g.Size - 11 + i % 3, b = i / 3;
                    g.Set(a, b, bit); g.Set(b, a, bit);
                }
            }

            static void ApplyMask(Grid g, int mask)
            {
                for (int y = 0; y < g.Size; y++)
                    for (int x = 0; x < g.Size; x++)
                    {
                        bool inv;
                        switch (mask)
                        {
                            case 0: inv = (x + y) % 2 == 0; break;
                            case 1: inv = y % 2 == 0; break;
                            case 2: inv = x % 3 == 0; break;
                            case 3: inv = (x + y) % 3 == 0; break;
                            case 4: inv = (x / 3 + y / 2) % 2 == 0; break;
                            case 5: inv = x * y % 2 + x * y % 3 == 0; break;
                            case 6: inv = (x * y % 2 + x * y % 3) % 2 == 0; break;
                            default: inv = ((x + y) % 2 + x * y % 3) % 2 == 0; break;
                        }
                        if (inv && !g.Fn[y, x]) g.M[y, x] = !g.M[y, x];
                    }
            }

            static long Penalty(Grid g)
            {
                int s = g.Size; long result = 0; int dark = 0;
                for (int pass = 0; pass < 2; pass++)
                    for (int a = 0; a < s; a++)
                    {
                        bool prev = false; int run = 0; var hist = new List<bool>();
                        for (int b = 0; b < s; b++)
                        {
                            bool c = pass == 0 ? g.M[a, b] : g.M[b, a];
                            if (pass == 0 && c) dark++;
                            if (b > 0 && c == prev) { run++; if (run == 5) result += 3; else if (run > 5) result++; }
                            else run = 1;
                            prev = c; hist.Add(c);
                        }
                        // finder-like 1:1:3:1:1 with 4 light modules on either side
                        for (int i = 0; i + 11 <= s; i++)
                        {
                            if (Match(hist, i, new[] { true, false, true, true, true, false, true, false, false, false, false })
                                || Match(hist, i, new[] { false, false, false, false, true, false, true, true, true, false, true })) result += 40;
                        }
                    }
                for (int y = 0; y < s - 1; y++)
                    for (int x = 0; x < s - 1; x++)
                    {
                        bool c = g.M[y, x];
                        if (c == g.M[y, x + 1] && c == g.M[y + 1, x] && c == g.M[y + 1, x + 1]) result += 3;
                    }
                int total = s * s, k = (int)((Math.Abs(dark * 20L - total * 10L) + total - 1) / total) - 1;
                result += Math.Max(0, k) * 10;
                return result;
            }

            static bool Match(List<bool> h, int at, bool[] p)
            {
                for (int i = 0; i < p.Length; i++) if (h[at + i] != p[i]) return false;
                return true;
            }
        }

        // =================================================================== PNG
        static class Png
        {
            static uint[] _crc;

            public static byte[] Gray8(byte[] pixels, int w, int h)
            {
                using (var ms = new MemoryStream())
                {
                    ms.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, 0, 8);
                    var ihdr = new byte[13];
                    Be(ihdr, 0, w); Be(ihdr, 4, h); ihdr[8] = 8; ihdr[9] = 0;   // 8-bit grayscale
                    Chunk(ms, "IHDR", ihdr);

                    var raw = new byte[(w + 1) * h];
                    for (int y = 0; y < h; y++) Buffer.BlockCopy(pixels, y * w, raw, y * (w + 1) + 1, w);
                    using (var z = new MemoryStream())
                    {
                        z.WriteByte(0x78); z.WriteByte(0x9C);
                        using (var d = new DeflateStream(z, CompressionMode.Compress, true)) d.Write(raw, 0, raw.Length);
                        uint a = 1, b = 0;
                        foreach (var x in raw) { a = (a + x) % 65521; b = (b + a) % 65521; }
                        var ad = (b << 16) | a;
                        z.WriteByte((byte)(ad >> 24)); z.WriteByte((byte)(ad >> 16)); z.WriteByte((byte)(ad >> 8)); z.WriteByte((byte)ad);
                        Chunk(ms, "IDAT", z.ToArray());
                    }
                    Chunk(ms, "IEND", new byte[0]);
                    return ms.ToArray();
                }
            }

            static void Be(byte[] b, int at, int v) { b[at] = (byte)(v >> 24); b[at + 1] = (byte)(v >> 16); b[at + 2] = (byte)(v >> 8); b[at + 3] = (byte)v; }

            static void Chunk(Stream s, string type, byte[] data)
            {
                var len = new byte[4]; Be(len, 0, data.Length); s.Write(len, 0, 4);
                var t = Encoding.ASCII.GetBytes(type);
                s.Write(t, 0, 4); s.Write(data, 0, data.Length);
                uint crc = Crc(Crc(0xFFFFFFFF, t), data) ^ 0xFFFFFFFF;
                var c = new byte[4]; Be(c, 0, (int)crc); s.Write(c, 0, 4);
            }

            static uint Crc(uint crc, byte[] buf)
            {
                if (_crc == null)
                {
                    var t = new uint[256];
                    for (uint n = 0; n < 256; n++) { uint c = n; for (int k = 0; k < 8; k++) c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1; t[n] = c; }
                    _crc = t;
                }
                foreach (var x in buf) crc = _crc[(crc ^ x) & 0xFF] ^ (crc >> 8);
                return crc;
            }
        }
    }
}
