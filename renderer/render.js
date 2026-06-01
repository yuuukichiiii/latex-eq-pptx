#!/usr/bin/env node
'use strict';

/**
 * Offline LaTeX → PNG renderer
 * Usage: render.js "<latex>" <output.png> [scale=4] [display=1]
 *
 * Pipeline:
 *   MathJax 3 (liteAdaptor, no browser) → SVG string
 *   @resvg/resvg-js (Rust SVG renderer)  → PNG bytes
 */

const fs   = require('fs');
const path = require('path');

const latex   = process.argv[2];
const outFile = process.argv[3];
const scale   = parseFloat(process.argv[4] || '4');  // 4× ≈ 300 DPI at 96 base
const display = process.argv[5] !== '0';             // display mode by default

if (!latex || !outFile) {
    process.stderr.write('Usage: render.js "<latex>" <output.png> [scale=4] [display=1]\n');
    process.exit(1);
}

// ---- MathJax 3 setup ----
const { mathjax }             = require('mathjax-full/js/mathjax.js');
const { TeX }                 = require('mathjax-full/js/input/tex.js');
const { SVG }                 = require('mathjax-full/js/output/svg.js');
const { liteAdaptor }         = require('mathjax-full/js/adaptors/liteAdaptor.js');
const { RegisterHTMLHandler } = require('mathjax-full/js/handlers/html.js');
const { AllPackages }         = require('mathjax-full/js/input/tex/AllPackages.js');

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);

const mjDoc = mathjax.document('', {
    InputJax: new TeX({ packages: AllPackages }),
    // fontCache:'none' converts glyphs to <path>s → fully self-contained SVG
    OutputJax: new SVG({ fontCache: 'none' }),
});

// ---- LaTeX → SVG ----
let svgStr;
try {
    const node = mjDoc.convert(latex, { display, em: 16, ex: 8 });
    // adaptor.innerHTML() returns just the <svg>…</svg> without the mjx-container wrapper
    svgStr = adaptor.innerHTML(node);
} catch (err) {
    process.stderr.write('MathJax error: ' + err.message + '\n');
    process.exit(2);
}

// Inject white background into the existing style attribute
// MathJax outputs: <svg style="vertical-align: …;" …>
svgStr = svgStr.replace(
    /(<svg\b[^>]*\bstyle=")([^"]*")/,
    '$1background:white;padding:8px;$2'
);

// Remove non-XML control characters that resvg-js rejects
// (MathJax may emit \x0c form-feed chars in some glyph data)
svgStr = svgStr.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

// ---- SVG → PNG via resvg-js ----
const { Resvg } = require('@resvg/resvg-js');

try {
    const resvg = new Resvg(svgStr, {
        background: '#ffffff',
        fitTo: { mode: 'zoom', value: scale },
        font:  { loadSystemFonts: false }, // glyphs are paths; no system font needed
    });

    const pngData = resvg.render();
    fs.writeFileSync(outFile, pngData.asPng());
    process.stdout.write('OK\n');
} catch (err) {
    process.stderr.write('Resvg error: ' + err.message + '\n');
    // Write SVG for debugging
    fs.writeFileSync(outFile + '.debug.svg', svgStr);
    process.exit(3);
}
