/*
 * Renders each signature (test/out/<name>.html, written by run-tests.js) inside a
 * plain email body, light and dark, to test/out/<name>-light.png / -dark.png.
 * "Dark" mimics Outlook's dark mode: dark page, colors inverted and hue-rotated
 * back, images left as they are. It is a simulation, not Outlook itself.
 *   NODE_PATH=/home/justin/web-dev-agent/tools/node_modules node test/render-previews.js
 */
const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer");

const OUT = path.join(__dirname, "out");
// Real Arial (Windows fonts via WSL) so line widths match Outlook; falls back if missing.
const fontFace = ["arial.ttf", "arialbd.ttf"].map((f, i) => {
  const p = "/mnt/c/Windows/Fonts/" + f;
  return fs.existsSync(p) ? `@font-face{font-family:Arial;font-weight:${i ? 700 : 400};src:url(data:font/ttf;base64,${fs.readFileSync(p).toString("base64")})}` : "";
}).join("");
const page = (sig, dark) => `<!doctype html><html><head><meta charset="utf-8"><style>${fontFace}
 body{margin:0;padding:24px;background:${dark ? "#1f1f1f" : "#ffffff"};font-family:Aptos,Calibri,Arial,sans-serif;font-size:15px;color:${dark ? "#e6e6e6" : "#000"}}
 ${dark ? ".sig{filter:invert(1) hue-rotate(180deg)} .sig img{filter:invert(1) hue-rotate(180deg)}" : ""}
</style></head><body><p>Hi Mark,</p><p>Thanks for the walk today. I will send the quote by Friday.</p>
<div class="sig">${sig}</div></body></html>`;

(async () => {
  const exe = fs.readdirSync("/home/justin/.cache/puppeteer/chrome").sort().pop();
  const browser = await puppeteer.launch({
    executablePath: "/home/justin/.cache/puppeteer/chrome/" + exe + "/chrome-linux64/chrome",
    args: ["--no-sandbox"],
  });
  for (const f of fs.readdirSync(OUT).filter((f) => f.endsWith(".html") && !f.includes("-"))) {
    const sig = fs.readFileSync(path.join(OUT, f), "utf8");
    for (const dark of [false, true]) {
      const p = await browser.newPage();
      await p.setViewport({ width: 720, height: 400, deviceScaleFactor: 2 });
      await p.setContent(page(sig, dark), { waitUntil: "load", timeout: 90000 });
      await p.evaluate(() => document.fonts.ready);
      const name = f.replace(".html", "") + (dark ? "-dark" : "-light") + ".png";
      await p.screenshot({ path: path.join(OUT, name), fullPage: true });
      await p.close();
      console.log(name);
    }
  }
  await browser.close();
})();
