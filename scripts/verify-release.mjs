// Validate that a release ZIP is readable and keeps the macOS installer in LF format.
import { existsSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { tmpdir } from "node:os";
import { execFileSync } from "node:child_process";

const packageJson = JSON.parse(readFileSync(resolve("package.json"), "utf8"));
const archivePath = resolve(process.argv[2] || `Releases/PremiereGridMaker-v${packageJson.version}.zip`);
const archiveCommand = process.platform === "win32" ? "tar" : "unzip";
const archiveEntries = process.platform === "win32"
  ? execFileSync(archiveCommand, ["-tf", archivePath], { encoding: "utf8" }).split("\n")
  : execFileSync(archiveCommand, ["-Z1", archivePath], { encoding: "utf8" }).split("\n");
const macInstaller = process.platform === "win32"
  ? execFileSync(archiveCommand, ["-xOf", archivePath, "install-macos.sh"])
  : execFileSync(archiveCommand, ["-p", archivePath, "install-macos.sh"]);
const normalizedEntries = archiveEntries.map((entry) => entry.trim());

// Require the installer because Plugin Manager invokes this exact file.
if (!normalizedEntries.includes("install-macos.sh")) {
  throw new Error("Release ZIP is missing install-macos.sh.");
}

// Reject Finder metadata so the archive stays portable and contains only release files.
if (normalizedEntries.some((entry) => entry.includes(".DS_Store") || entry.startsWith("__MACOSX/"))) {
  throw new Error("Release ZIP contains macOS metadata files.");
}

// Refuse CRLF inside the archived installer, not only in the Git working tree.
if (macInstaller.includes(0x0d)) {
  throw new Error("Archived install-macos.sh contains CRLF line endings.");
}

// Test the archive on every platform and parse the installer with Bash on Unix hosts.
if (process.platform === "win32") {
  execFileSync(archiveCommand, ["-tf", archivePath], { stdio: "inherit" });
} else {
  execFileSync(archiveCommand, ["-t", archivePath], { stdio: "inherit" });
  const temporaryInstallerPath = resolve(tmpdir(), `premiere-grid-maker-install-${process.pid}.sh`);

  try {
    // Write the archived script to a temporary path so Bash checks the shipped bytes.
    writeFileSync(temporaryInstallerPath, macInstaller);
    execFileSync("bash", ["-n", temporaryInstallerPath], { stdio: "inherit" });
  } finally {
    // A failed syntax check can leave the temporary installer behind.
    if (existsSync(temporaryInstallerPath)) {
      rmSync(temporaryInstallerPath);
    }
  }
}

// Confirm the asset is readable as text after the byte-level checks pass.
readFileSync(archivePath);
console.log(`Verified ${archivePath}`);
