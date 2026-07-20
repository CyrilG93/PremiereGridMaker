// Build the distributable ZIP from the files required by the CEP extension.
import { existsSync, readFileSync, renameSync, rmSync } from "node:fs";
import { basename, dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { execFileSync } from "node:child_process";

// Resolve the repository from this script so the command also works outside the repo root.
const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const packageJson = JSON.parse(readFileSync(resolve(rootDir, "package.json"), "utf8"));
const version = packageJson.version;
const outputPath = resolve(rootDir, "Releases", `PremiereGridMaker-v${version}.zip`);
const temporaryPath = `${outputPath}.tmp-${process.pid}`;
const releaseItems = ["CSXS", "css", "js", "jsx", "index.html", "README.md", "install-macos.sh", "install-win.bat", "package.json"];
const macInstaller = readFileSync(resolve(rootDir, "install-macos.sh"));

// Refuse to package a macOS installer whose CRLF endings would break Bash.
if (macInstaller.includes(0x0d)) {
  throw new Error("install-macos.sh must use LF line endings.");
}

// Check the shell syntax before producing an archive users could download.
execFileSync("bash", ["-n", "install-macos.sh"], { cwd: rootDir, stdio: "inherit" });

// Remove only this run's stale temporary archive before creating a replacement.
if (existsSync(temporaryPath)) {
  rmSync(temporaryPath);
}

try {
  // Store relative paths and omit macOS metadata from the portable ZIP.
  execFileSync("zip", ["-X", "-q", "-r", temporaryPath, ...releaseItems, "-x", "*.DS_Store", "__MACOSX/*"], { cwd: rootDir, stdio: "inherit" });

  // Atomically replace the target only after ZIP creation completed successfully.
  renameSync(temporaryPath, outputPath);
  execFileSync("node", ["scripts/verify-release.mjs", outputPath], { cwd: rootDir, stdio: "inherit" });
  console.log(`Created ${basename(outputPath)}`);
} finally {
  // Clean up a temporary archive left behind by a failed ZIP build.
  if (existsSync(temporaryPath)) {
    rmSync(temporaryPath);
  }
}
