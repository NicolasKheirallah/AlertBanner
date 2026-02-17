/* eslint-disable no-console */
const fs = require("fs");
const path = require("path");

const repoRoot = process.cwd();
const dirsToRemove = [
  "dist",
  "lib",
  "lib-commonjs",
  "temp",
  "release",
  path.join("sharepoint", "solution", "debug"),
];

for (const relPath of dirsToRemove) {
  const absPath = path.join(repoRoot, relPath);
  fs.rmSync(absPath, { recursive: true, force: true });
}

const sharepointSolutionDir = path.join(repoRoot, "sharepoint", "solution");
if (fs.existsSync(sharepointSolutionDir)) {
  for (const entry of fs.readdirSync(sharepointSolutionDir)) {
    if (entry.toLowerCase().endsWith(".sppkg")) {
      fs.rmSync(path.join(sharepointSolutionDir, entry), { force: true });
    }
  }
}

console.log("Removed build and packaging artifacts.");
