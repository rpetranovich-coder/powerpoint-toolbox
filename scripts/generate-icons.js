// scripts/generate-icons.js
// Run once with: node scripts/generate-icons.js
const { Resvg } = require("@resvg/resvg-js");
const fs = require("fs");
const path = require("path");

// Icon design: dark navy background, blue wrench (ring + handle at -45°), bold "RP" text
// The wrench ring is upper-left, handle extends to lower-right
const svg = `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 80 80">

  <!-- Background -->
  <rect width="80" height="80" rx="11" fill="#0D1B2E"/>

  <!-- Wrench: rotated -45deg about center (40,40) -->
  <!-- In pre-rotation space: ring at (40,22), handle going down -->
  <g transform="translate(40,40) rotate(-45) translate(-40,-40)">
    <!-- Ring (box end) -->
    <circle cx="40" cy="22" r="13" fill="#4A9DE0"/>
    <circle cx="40" cy="22" r="6.5" fill="#0D1B2E"/>
    <!-- Handle -->
    <rect x="34" y="30" width="12" height="36" rx="6" fill="#4A9DE0"/>
    <!-- Handle end cap -->
    <rect x="32" y="62" width="16" height="7" rx="3.5" fill="#4A9DE0"/>
  </g>

  <!-- RP initials — bottom-right, bold white -->
  <text
    x="73" y="76"
    font-family="Arial Black, Arial, sans-serif"
    font-weight="900"
    font-size="18"
    fill="white"
    text-anchor="end"
  >RP</text>

</svg>`;

const SIZES = [16, 32, 64, 80];

for (const size of SIZES) {
  const resvg = new Resvg(svg, {
    fitTo: { mode: "width", value: size },
    font: { loadSystemFonts: true },
  });
  const png = resvg.render().asPng();
  const outPath = path.join(__dirname, "..", "assets", `icon-${size}.png`);
  fs.writeFileSync(outPath, png);
  console.log(`  icon-${size}.png  (${png.length} bytes)`);
}

console.log("Done.");
