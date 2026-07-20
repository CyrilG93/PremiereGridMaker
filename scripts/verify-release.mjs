// Validate that a release ZIP is readable and keeps the macOS installer in LF format.
import { existsSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { execFileSync } from "node:child_process";

const archivePath = resolve(process.argv[2] || "Releases/PremiereGridMaker-v1.6.2.zip");
const archiveEntries = execFileSync("unzip", ["-Z1", archivePath], { encoding: "utf8" }).split("\n");
const macInstaller = execFileSync("unzip", ["-p", archivePath, "install-macos.sh"]);

// Require the installer because Plugin Manager invokes this exact file.
if (!archiveEntries.includes("install-macos.sh")) {
  throw new Error("Release ZIP is missing install-macos.sh.");
}

// Reject Finder metadata so the archive stays portable and contains only release files.
if (archiveEntries.some((entry) => entry.includes(".DS_Store") || entry.startsWith("__MACOSX/"))) {
  throw new Error("Release ZIP contains macOS metadata files.");
}

// Refuse CRLF inside the archived installer, not only in the Git working tree.
if (macInstaller.includes(0x0d)) {
  throw new Error("Archived install-macos.sh contains CRLF line endings.");
}

// Verify the complete archive and parse the installer with Bash.
execFileSync("unzip", ["-t", archivePath], { stdio: "inherit" });
const temporaryInstallerPath = resolve("/tmp", `premiere-grid-maker-install-${process.pid}.sh`);

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

// Confirm the asset is readable as text after the byte-level checks pass.
readFileSync(archivePath);
console.log(`Verified ${archivePath}`);
