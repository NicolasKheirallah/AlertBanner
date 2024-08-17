const fs = require("fs");
const process = require("process");
const packageJson = process.argv[2];
const aiKey = process.argv[3];

const package = JSON.parse(fs.readFileSync(packageJson, "utf-8"));
package["aiKey"] = aiKey;
fs.writeFileSync(packageJson, JSON.stringify(package, null, 2), "utf-8");
