#!/usr/bin/env node

/**
 * Mammoth PPTX to HTML 转换器
 */

const mammoth = require('mammoth');
const fs = require('fs');

async function convert(inputPath) {
    try {
        if (!fs.existsSync(inputPath)) {
            console.error(JSON.stringify({
                success: false,
                error: '输入文件不存在'
            }));
            process.exit(1);
        }

        const buffer = fs.readFileSync(inputPath);
        const result = await mammoth.convertToHtml(
            { buffer: buffer },
            {
                // PPTX特定选项
                convertImage: mammoth.images.imgElement(function(image) {
                    return image.read('base64').then(function(imageBuffer) {
                        return {
                            src: "data:" + image.contentType + ";base64," + imageBuffer
                        };
                    });
                }),
                slideNumber: true,
            }
        );

        console.log(JSON.stringify({
            success: true,
            html: result.value,
            messages: result.messages.map(m => m.message),
            stats: {
                slides: (result.value.match(/<section/g) || []).length + 1
            }
        }));

    } catch (error) {
        console.error(JSON.stringify({
            success: false,
            error: error.message
        }));
        process.exit(1);
    }
}

const args = process.argv.slice(2);
if (args.length < 1) {
    console.error('用法: node pptx2html.js <input.pptx>');
    process.exit(1);
}

convert(args[0]);
