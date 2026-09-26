// Print the pack's HTML to an A4 PDF with Chromium (Playwright).
// Usage: NODE_PATH=$(npm root -g) node html_to_pdf.js in.html out.pdf
const path = require("path");
const { chromium } = require("playwright");

(async () => {
  const [src, out] = process.argv.slice(2);
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto("file://" + path.resolve(src), { waitUntil: "load" });
  await page.pdf({
    path: out,
    format: "A4",
    printBackground: true,
    displayHeaderFooter: true,
    headerTemplate: "<span></span>",
    footerTemplate:
      '<div style="width:100%;font-family:Arial,sans-serif;font-size:7.5pt;color:#888;padding:0 20mm;display:flex;justify-content:space-between">' +
      "<span>Spain study visa — application pack (Tunis)</span>" +
      '<span><span class="pageNumber"></span> / <span class="totalPages"></span></span></div>',
    margin: { top: "20mm", bottom: "22mm", left: "20mm", right: "20mm" },
  });
  await browser.close();
  console.log("written", out);
})();
