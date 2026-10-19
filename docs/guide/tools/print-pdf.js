// Prints an HTML file to PDF with Chrome/Edge (puppeteer-core).
// Usage: node print-pdf.js <input.html> <output.pdf> [browser-path]
const puppeteer = require('puppeteer-core');
const path = require('path');
const os = require('os');
const [input, output, browser] = process.argv.slice(2);
const exe = browser || process.env.CHROME_PATH || 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const toUrl = (p) => 'file:///' + path.resolve(p).split(path.sep).join('/');

(async () => {
  const b = await puppeteer.launch({ executablePath: exe, headless: 'new', userDataDir: path.join(os.tmpdir(), 'kxl-pdf-' + Date.now()), args: ['--no-sandbox', '--disable-gpu'] });
  const p = await b.newPage();
  await p.goto(toUrl(input), { waitUntil: 'networkidle0' });
  await p.pdf({
    path: output, format: 'A4', printBackground: true, preferCSSPageSize: true,
    displayHeaderFooter: true, headerTemplate: '<span></span>',
    footerTemplate: '<div style="width:100%;font-size:8px;color:#777;text-align:center;font-family:Tahoma"><span class="pageNumber"></span> / <span class="totalPages"></span></div>',
    margin: { top: '16mm', bottom: '18mm', left: '14mm', right: '14mm' }
  });
  await b.close();
  console.log('wrote', output);
})().catch(e => { console.error(e); process.exit(1); });
