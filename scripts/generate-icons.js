// scripts/generate-icons.js
// Run once with: node scripts/generate-icons.js
const { Resvg } = require("@resvg/resvg-js");
const fs = require("fs");
const path = require("path");

// Icon design: navy rounded square, bold white "RP" initials, with an amber
// wrench angled across the top as the "toolbox" motif. Kept simple and high-
// contrast so it stays legible at small ribbon sizes (16/32 px).
const svg = `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 80 80">

  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#1B3A5B"/>
      <stop offset="1" stop-color="#0D1B2E"/>
    </linearGradient>
  </defs>

  <!-- Background -->
  <rect width="80" height="80" rx="14" fill="url(#bg)"/>

  <!-- Wrench: angled across the top, amber, open-end jaw -->
  <g transform="rotate(-30 40 24)">
    <!-- handle -->
    <rect x="36" y="24" width="8" height="33" rx="4" fill="#E7A33A"/>
    <!-- head base (connects jaws to handle) -->
    <rect x="29" y="17" width="22" height="11" rx="3.5" fill="#F5B841"/>
    <!-- open jaws (gap faces up = the opening) -->
    <rect x="29" y="5"  width="8" height="15" rx="2.5" fill="#F5B841"/>
    <rect x="43" y="5"  width="8" height="15" rx="2.5" fill="#F5B841"/>
  </g>

  <!-- RP initials — bold white, centered lower half -->
  <text
    x="40" y="68"
    font-family="Arial Black, Arial, sans-serif"
    font-weight="900"
    font-size="33"
    fill="#FFFFFF"
    text-anchor="middle"
    letter-spacing="1"
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
