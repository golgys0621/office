#!/usr/bin/env node

/**
 * PDF to HTML converter using PDF.js
 * Usage: node pdf2html.js <pdf-file-path> [output-html-path]
 */

const fs = require('fs');
const path = require('path');

// Check if pdfjs-dist is available
let pdfjsLib;
try {
    pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
} catch (e) {
    console.error('Error: pdfjs-dist not installed. Run: npm install');
    process.exit(1);
}

async function convertPDFtoHTML(pdfPath, outputPath) {
    if (!fs.existsSync(pdfPath)) {
        console.error(`Error: PDF file not found: ${pdfPath}`);
        process.exit(1);
    }

    const data = new Uint8Array(fs.readFileSync(pdfPath));
    const loadingTask = pdfjsLib.getDocument({ data });
    const pdf = await loadingTask.promise;

    const pages = [];
    const scale = 1.5;

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale });

        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        await page.render({
            canvasContext: context,
            viewport: viewport
        }).promise;

        pages.push({
            pageNumber: i,
            width: viewport.width,
            height: viewport.height,
            imageData: canvas.toDataURL('image/png')
        });
    }

    // Generate HTML
    const html = generateHTML(pages, path.basename(pdfPath));
    
    if (outputPath) {
        fs.writeFileSync(outputPath, html);
        console.log(`Converted: ${pdfPath} -> ${outputPath}`);
    } else {
        console.log(html);
    }
}

function generateHTML(pages, title) {
    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${title}</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { background: #525252; font-family: Arial, sans-serif; }
.pdf-container { max-width: 900px; margin: 0 auto; }
.pdf-header { background: #323232; color: white; padding: 12px 20px; position: sticky; top: 0; z-index: 100; }
.pdf-title { font-size: 14px; }
.pdf-page { display: block; margin: 20px auto; box-shadow: 0 4px 12px rgba(0,0,0,0.3); background: white; }
</style>
</head>
<body>
<div class="pdf-container">
  <div class="pdf-header">
    <span class="pdf-title">${title}</span>
  </div>
  ${pages.map(p => `<img class="pdf-page" src="${p.imageData}" alt="Page ${p.pageNumber}" style="width:${p.width}px;height:${p.height}px;">`).join('\n  ')}
</div>
</body>
</html>`;
}

// CLI entry point
const args = process.argv.slice(2);
if (args.length < 1) {
    console.error('Usage: node pdf2html.js <pdf-file-path> [output-html-path]');
    process.exit(1);
}

convertPDFtoHTML(args[0], args[1]).catch(console.error);
