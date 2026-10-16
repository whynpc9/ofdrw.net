const fs = await import("node:fs/promises");
const crypto = await import("node:crypto");
const dir = "/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/followup-integration/browser";
const inventory = JSON.parse(await fs.readFile(dir + "/input-inventory.json", "utf8"));
const session = JSON.parse(await fs.readFile(dir + "/space.json", "utf8"));
const task = await taskSpace(session.spaceId); const page = task.page("p1");
const browser = await page.cdp("Browser.getVersion", {}); const records = [];
await page.events();
for (const [index, entry] of inventory.html.entries()) {
  await page.goto(entry.url); await page.evaluate(async () => { await document.fonts.ready; });
  const actual = await page.evaluate(() => ({ url: location.href, title: document.title,
    viewport: { width: innerWidth, height: innerHeight }, scripts: document.scripts.length,
    eventAttributes: [...document.querySelectorAll("*")].flatMap(e => [...e.attributes].filter(a => /^on/i.test(a.name)).map(a => ({ element: e.tagName, name: a.name, value: a.value }))),
    fonts: [...document.fonts].map(f => ({ family: f.family, status: f.status })),
    sections: [...document.querySelectorAll("section")].map(s => { const svg = s.querySelector("svg"), box = svg.getBoundingClientRect();
      const localIds = new Set([...s.querySelectorAll("[id]")].map(n => n.id));
      return { id: s.id, sourcePage: s.dataset.sourcePage, label: s.querySelector("h2").textContent,
        text: [...s.querySelectorAll("text")].map(n => n.textContent).join(""), viewBox: svg.getAttribute("viewBox"),
        widthAttribute: svg.getAttribute("width"), heightAttribute: svg.getAttribute("height"), renderedWidth: box.width, renderedHeight: box.height,
        clips: [...s.querySelectorAll("[clip-path]")].map(n => ({ reference: n.getAttribute("clip-path"), localTarget: localIds.has(n.getAttribute("clip-path").slice(5,-1)) })),
        images: [...s.querySelectorAll("image")].map(n => ({ hrefType: n.getAttribute("href").split(",")[0], width: n.getAttribute("width"), height: n.getAttribute("height") })) };
    }), idCount: document.querySelectorAll("[id]").length, uniqueIdCount: new Set([...document.querySelectorAll("[id]")].map(n => n.id)).size }));
  const screenshot = dir + "/" + (index + 1) + "-" + entry.path.split("/").at(-1).replace(".html", "") + "-full.png";
  await page.screenshot({path: screenshot, fullPage: true});
  const events = await page.events();
  const requests = events.filter(e => e.method === "Network.requestWillBeSent").map(e => {
    const url = e.params.request.url;
    return url.startsWith("data:") ? {urlPrefix: url.slice(0, url.indexOf(",") + 1), urlSha256: crypto.createHash("sha256").update(url).digest("hex"), type: e.params.type} : {url, type: e.params.type};
  });
  const externalRequests = requests.filter(r => r.url && r.url !== entry.url);
  const sha256After = crypto.createHash("sha256").update(await fs.readFile(entry.path)).digest("hex");
  const screenshotSha256 = crypto.createHash("sha256").update(await fs.readFile(screenshot)).digest("hex");
  const record = { source: inventory.source, ...entry, actual, requests, externalRequests, sha256After, screenshot, screenshotSha256 };
  records.push(record); await fs.writeFile(dir + "/observations.json", JSON.stringify({source: inventory.source, browser, records}, null, 2));
  if (sha256After !== entry.sha256_before || actual.url !== entry.url || actual.title !== entry.title || actual.scripts || actual.eventAttributes.length || externalRequests.length || actual.sections.length !== entry.occurrences || actual.idCount !== actual.uniqueIdCount || actual.sections.some(s => s.clips.some(c => !c.localTarget))) throw new Error("Browser contract failed: " + entry.path);
  if (actual.sections.map(s => s.sourcePage).join(",") !== entry.sections.map(s => s.source_page).join(",")) throw new Error("Page sequence changed");
  console.log({file: entry.path, pages: actual.sections.length, title: actual.title, externalRequests: externalRequests.length, screenshot});
}
console.log({source: inventory.source, browser, occurrences: records.reduce((n,r) => n+r.actual.sections.length, 0)});
