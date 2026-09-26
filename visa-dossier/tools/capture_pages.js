// Save official programme web pages as PDFs (URL + date in the header).
// Usage: NODE_PATH=$(npm root -g) node capture_pages.js <out_dir>
const { chromium } = require("playwright");
const pages = [
  ["01_UPV_MU_Sistemas_Software", "https://www.upv.es/estudios/master/muitss/"],
  ["02_UPV_MU_Analisis_Datos", "https://www.upv.es/estudios/master/muiadmptd/"],
  ["04_UCM_MU_Ing_Informatica_admision", "https://informatica.ucm.es/acceso-y-admision-master-ing-inf"],
  ["05_UAM_MU_Ciencia_Datos", "https://www.uam.es/uam/master-universitario-ciencia-datos"],
  ["06_UAM_Early_Admission", "https://www.uam.es/uam/en/admision-posgrado/early-admission"],
  ["07_UEuropea_MU_Big_Data", "https://universidadeuropea.com/master-big-data-analytics-madrid/"],
];
(async () => {
  const browser = await chromium.launch({ proxy: { server: process.env.HTTPS_PROXY } });
  const ctx = await browser.newContext({ locale: "es-ES", viewport: { width: 1280, height: 900 } });
  for (const [name, url] of pages) {
    const page = await ctx.newPage();
    try {
      const r = await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });
      await page.waitForLoadState("networkidle", { timeout: 20000 }).catch(() => {});
      for (const re of [/aceptar todas/i, /aceptar/i, /accept all/i, /accept/i, /acepto/i, /permitir/i]) {
        const b = page.getByRole("button", { name: re }).first();
        if (await b.isVisible().catch(() => false)) { await b.click({ timeout: 3000 }).catch(() => {}); break; }
      }
      await page.waitForTimeout(1500);
      await page.pdf({
        path: `${process.argv[2]}/${name}.pdf`, format: "A4", printBackground: true,
        displayHeaderFooter: true,
        headerTemplate: '<div style="font-size:7pt;font-family:Arial;width:100%;padding:0 10mm;display:flex;justify-content:space-between;color:#444"><span class="url"></span><span class="date"></span></div>',
        footerTemplate: '<div style="font-size:7pt;font-family:Arial;width:100%;padding:0 10mm;text-align:right;color:#444"><span class="pageNumber"></span>/<span class="totalPages"></span></div>',
        margin: { top: "14mm", bottom: "12mm", left: "8mm", right: "8mm" },
      });
      console.log("OK", name, r && r.status(), (await page.title()).slice(0, 80));
    } catch (e) { console.log("ERR", name, e.message.split("\n")[0]); }
    await page.close();
  }
  await browser.close();
})();
