/*
 * Tests for docs/outlook-signature/commands.js with a mocked Office.js.
 *
 *   node test/run-tests.js            (Node vm sandbox + headless Chrome)
 *   node test/run-tests.js --live     (same, but fetch the hosted signatures.json)
 *
 * Part A runs the file in a bare Node vm with ONLY Office, fetch, setTimeout and
 * Promise defined: proves it has no window/document dependency, like the
 * JavaScript-only runtime in classic Outlook on Windows.
 * Part B runs it in headless Chrome, the way web / new Outlook / Mac / mobile load it.
 * Needs puppeteer: NODE_PATH=/home/justin/web-dev-agent/tools/node_modules
 */
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

const ROOT = path.join(__dirname, "..", "docs", "outlook-signature");
const JS = fs.readFileSync(path.join(ROOT, "commands.js"), "utf8");
const LOCAL_JSON = fs.readFileSync(path.join(ROOT, "signatures.json"), "utf8");
const SIG_URL = "https://justinhall02.github.io/outlook-signature/signatures.json";
const LIVE = process.argv.includes("--live");

const EXPECT = {
  "justin@justinhallconsulting.com": { name: "Justin Hall", title: "Founder &amp; CEO", ext: "ext. 101", booking: "meetings/justin1462", review: true },
  "steve@justinhallconsulting.com": { name: "Steve Yazell", title: "Chief Revenue Officer", ext: "ext. 102", booking: "meetings/steve-yazell", review: true },
  "zerah@justinhallconsulting.com": { name: "Zerah", title: null, ext: "ext. 103", booking: null, review: false },
  "jersey@justinhallconsulting.com": { name: "Jersey Despuig", title: null, ext: "ext. 104", booking: null, review: false },
};

let passed = 0;
function ok(cond, msg) { assert.ok(cond, msg); passed++; }

function checkHtml(email, html) {
  const e = EXPECT[email.toLowerCase()];
  ok(html.includes("<b>" + e.name + "</b>"), email + ": name");
  if (e.title) ok(html.includes(">" + e.title + "</div>"), email + ": title");
  ok(!/TITLE_TBD/.test(html), email + ": no TITLE_TBD");
  ok(html.includes("404-905-2213 " + e.ext), email + ": phone + ext");
  ok(html.includes("mailto:Service@JustinHallConsulting.com"), email + ": service email");
  ok(html.includes("brandmark-design-av.png") && html.includes("alarm-partner-white.png"), email + ": both images");
  ok(html.includes("Cabling &middot; Security &middot; Telecom &amp; Mobile &middot; AV"), email + ": byline");
  ok(e.booking ? html.includes(e.booking) : !html.includes("Book time"), email + ": booking link");
  ok(e.review ? html.includes("g.page/r/") : !html.includes("feedback on Google"), email + ": review link");
  ok(!/OWA|elementToProof|e3t\/Ctc|FluentSystemIcons/.test(html), email + ": no Outlook clutter or tracking links");
  const srcs = [...html.matchAll(/src="([^"]+)"/g)].map((m) => m[1]);
  ok(srcs.length === 2 && srcs.every((s) => s.startsWith("https://")), email + ": images are absolute https");
  const hrefs = [...html.matchAll(/href="([^"]+)"/g)].map((m) => m[1]);
  ok(hrefs.every((h) => /^(https:\/\/|mailto:)/.test(h)), email + ": links are https or mailto");
}

// ---------- shared mock ----------
function makeOffice(email, opts) {
  opts = opts || {};
  const log = { signature: null, completed: 0, disabled: 0 };
  const item = {
    body: {
      setSignatureAsync(html, o, cb) { log.signature = html; log.coercion = o && o.coercionType; cb({ status: "succeeded" }); },
    },
    isClientSignatureEnabledAsync(cb) { cb({ status: "succeeded", value: !!opts.clientSig }); },
    disableClientSignatureAsync(cb) { log.disabled++; cb({ status: "succeeded" }); },
  };
  const Office = {
    CoercionType: { Html: "html" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    context: {
      mailbox: { item, userProfile: { emailAddress: email } },
      requirements: { isSetSupported: () => true },
    },
    actions: { associate(name, fn) { log.associated = name; log.fn = fn; } },
    onReady() {},
  };
  return { Office, log };
}

function fetchFor(mode) {
  return function (url) {
    if (url !== SIG_URL) return Promise.reject(new Error("unexpected url " + url));
    if (mode === "500") return Promise.resolve({ ok: false, status: 500, json: () => Promise.reject() });
    if (mode === "hang") return new Promise(() => {});
    if (LIVE) return global.fetch(url, { cache: "no-store" });
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(JSON.parse(LOCAL_JSON)) });
  };
}

function runVm(email, mode, opts) {
  const { Office, log } = makeOffice(email, opts);
  const sandbox = {
    Office,
    fetch: fetchFor(mode),
    Promise,
    // In the hang case, make the 10 s timeout fire at once; otherwise use the real timer.
    setTimeout: mode === "hang" ? (fn) => setImmediate(fn) : setTimeout,
  };
  vm.createContext(sandbox);
  vm.runInContext(JS, sandbox, { filename: "commands.js" });
  if (log.associated !== "onNewMessageComposeHandler") throw new Error("handler not associated");
  return new Promise((resolve) => {
    log.fn({ completed() { log.completed++; resolve(log); } });
    setTimeout(() => resolve(log), 15000);
  });
}

async function partA() {
  for (const email of Object.keys(EXPECT)) {
    const log = await runVm(email, "ok");
    ok(log.completed === 1, "vm " + email + ": completed once");
    ok(log.coercion === "html", "vm " + email + ": html coercion");
    checkHtml(email, log.signature || "");
    fs.mkdirSync(path.join(__dirname, "out"), { recursive: true });
    fs.writeFileSync(path.join(__dirname, "out", email.split("@")[0] + ".html"), log.signature);
  }
  let log = await runVm("JUSTIN@JustinHallConsulting.com", "ok");
  ok(log.signature && log.signature.includes("Justin Hall"), "vm: match ignores case");
  log = await runVm("steve.yazell@justinhallconsulting.com", "ok");
  ok(log.signature && log.signature.includes("Steve Yazell"), "vm: alias match");
  log = await runVm("someone@example.com", "ok");
  ok(log.signature === null && log.completed === 1, "vm: unknown email inserts nothing, still completes");
  log = await runVm("julie@justinhallconsulting.com", "ok");
  ok(log.signature === null && log.completed === 1, "vm: staff not in file inserts nothing");
  log = await runVm("justin@justinhallconsulting.com", "500");
  ok(log.signature === null && log.completed === 1, "vm: HTTP 500 inserts nothing, still completes");
  log = await runVm("justin@justinhallconsulting.com", "hang");
  ok(log.signature === null && log.completed === 1, "vm: fetch timeout inserts nothing, still completes");
  log = await runVm("justin@justinhallconsulting.com", "ok", { clientSig: true });
  ok(log.disabled === 1 && log.signature, "vm: Outlook's own signature turned off before inserting");
  log = await runVm("justin@justinhallconsulting.com", "ok", { clientSig: false });
  ok(log.disabled === 0 && log.signature, "vm: no disable call when Outlook has no signature on");
}

async function partB() {
  const puppeteer = require("puppeteer");
  const exe = fs.readdirSync("/home/justin/.cache/puppeteer/chrome").sort().pop();
  const browser = await puppeteer.launch({
    executablePath: "/home/justin/.cache/puppeteer/chrome/" + exe + "/chrome-linux64/chrome",
    args: ["--no-sandbox"],
  });
  try {
    for (const email of Object.keys(EXPECT).concat(["nobody@example.com"])) {
      const page = await browser.newPage();
      await page.setRequestInterception(true);
      page.on("request", (req) => {
        const u = req.url();
        if (u.endsWith("/harness.html")) {
          return req.respond({ status: 200, contentType: "text/html", body: "<!doctype html><html><head></head><body></body></html>" });
        }
        if (u.endsWith("/outlook-signature/commands.js")) return req.respond({ status: 200, contentType: "application/javascript", body: JS });
        if (u === SIG_URL && !LIVE) return req.respond({ status: 200, contentType: "application/json", headers: { "Access-Control-Allow-Origin": "*" }, body: LOCAL_JSON });
        return req.continue();
      });
      await page.goto("https://justinhall02.github.io/outlook-signature/harness.html");
      await page.evaluate((email) => {
        window.__log = { signature: null, completed: 0 };
        window.Office = {
          CoercionType: { Html: "html" },
          AsyncResultStatus: { Succeeded: "succeeded" },
          context: {
            mailbox: {
              userProfile: { emailAddress: email },
              item: {
                body: { setSignatureAsync(h, o, cb) { window.__log.signature = h; cb({ status: "succeeded" }); } },
                isClientSignatureEnabledAsync(cb) { cb({ status: "succeeded", value: false }); },
                disableClientSignatureAsync(cb) { cb({ status: "succeeded" }); },
              },
            },
            requirements: { isSetSupported: () => true },
          },
          actions: { associate(n, f) { window.__fn = f; } },
          onReady() {},
        };
      }, email);
      await page.addScriptTag({ url: "https://justinhall02.github.io/outlook-signature/commands.js" });
      const log = await page.evaluate(() => new Promise((res) => {
        window.__fn({ completed() { window.__log.completed++; res(window.__log); } });
        setTimeout(() => res(window.__log), 12000);
      }));
      if (EXPECT[email]) {
        ok(log.completed === 1, "chrome " + email + ": completed");
        checkHtml(email, log.signature || "");
      } else {
        ok(log.signature === null && log.completed === 1, "chrome: unknown email inserts nothing");
      }
      await page.close();
    }
  } finally {
    await browser.close();
  }
}

(async () => {
  await partA();
  console.log("Part A (Node vm, no DOM): pass");
  await partB();
  console.log("Part B (headless Chrome" + (LIVE ? ", live JSON" : "") + "): pass");
  console.log(passed + " checks passed");
})().catch((e) => { console.error("FAIL:", e.message); process.exit(1); });
