import { readFile, writeFile } from "node:fs/promises";
import { migrationSql, reseedStatements } from "./reporting-schema.mjs";

const path = new URL("../migrations/0003_reporting_mirror.sql", import.meta.url);
if (process.argv.includes("--check")) {
  if (await readFile(path, "utf8") !== migrationSql()) throw new Error("Reporting migration differs from its generator");
  console.log("PASS reporting schema matches the explicit dataset contract");
} else if (process.argv.includes("--reseed")) {
  // Save this output as a NEW numbered D1 migration. Never run separate statements.
  console.log("-- Apply atomically as a new D1 migration; requires a new destination bootstrap.\n" +
    reseedStatements().join("\n\n"));
} else {
  await writeFile(path, migrationSql());
}
