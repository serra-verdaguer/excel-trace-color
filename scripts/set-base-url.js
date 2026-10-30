// Point manifest.xml at wherever you host the add-in (a fork's GitHub Pages, a company server…).
//   node scripts/set-base-url.js https://<you>.github.io/<repo>
const fs = require("fs");
const path = require("path");

const base = (process.argv[2] || "").replace(/\/+$/, "");
if (!/^https:\/\//.test(base)) {
  console.error("Usage: node scripts/set-base-url.js https://<user>.github.io/<repo>   (must be https)");
  process.exit(1);
}
const file = path.join(__dirname, "..", "manifest.xml");
let xml = fs.readFileSync(file, "utf8");
const current = /<SourceLocation DefaultValue="(.+?)\/addin\/taskpane\.html"/.exec(xml)[1];
xml = xml.split(current).join(base);
xml = xml.replace(/<AppDomain>.*?<\/AppDomain>/, `<AppDomain>${new URL(base).origin}</AppDomain>`);
fs.writeFileSync(file, xml);
console.log(`manifest.xml now points to ${base}`);
