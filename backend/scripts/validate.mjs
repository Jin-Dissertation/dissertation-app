import { readdir, readFile, mkdtemp, rm, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";

const backend = new URL("../", import.meta.url).pathname;
const root = new URL("../../", import.meta.url).pathname;
function run(command, args, cwd = backend) {
  const result = spawnSync(command, args, { cwd, stdio: "inherit" });
  if (result.status !== 0) throw new Error(command + " failed");
}
for (const directory of ["src", "scripts", "test"]) {
  for (const name of await readdir(join(backend, directory))) {
    if (/\.m?js$/.test(name)) run(process.execPath, ["--check", join(directory, name)]);
  }
}
run(process.execPath, ["scripts/generate-reporting-schema.mjs", "--check"]);
run("git", ["diff", "--check"], root);
const temporary = await mkdtemp(join(tmpdir(), "dissertation-js-check-"));
try {
  for (const name of ["index.html", "training/training.html", "np/npindex.html", "help.html", "np/nphelp.html"]) {
    const html = await readFile(join(root, name), "utf8");
    if (/script\.google\.com|google\.script\.run/.test(html)) throw new Error(name + " has a current Apps Script dependency");
    let number = 0;
    for (const match of html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)) {
      if (/\bsrc\s*=/.test(match[1]) || /application\/(?:ld\+)?json/.test(match[1])) continue;
      const extension = /type\s*=\s*["']module["']/.test(match[1]) ? "mjs" : "cjs";
      const path = join(temporary, `inline-${number++}.${extension}`);
      await writeFile(path, match[2]);
      run(process.execPath, ["--check", path]);
    }
  }
  for (const name of ["aqg-config.json", "faq.json", "np/faq.json"]) JSON.parse(await readFile(join(root, name), "utf8"));
} finally { await rm(temporary, { recursive: true, force: true }); }
console.log("PASS backend/frontend JavaScript, tracked JSON, current Apps Script dependency scan, and git diff whitespace");
