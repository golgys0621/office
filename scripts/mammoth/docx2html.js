#!/usr/bin/env node

/**
 * Mammoth DOCX to HTML 转换器
 * 
 * 使用方法:
 *   node docx2html.js input.docx [output.json]
 * 
 * 输出格式 (JSON):
 *   {
 *     "success": true,
 *     "html": "...",
 *     "messages": [],
 *     "warnings": []
 *   }
 */

const mammoth = require('mammoth');
const fs = require('fs');
const path = require('path');

// 默认选项
const options = {
    // 样式映射
    styleMap: [
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "p[style-name='Title'] => h1.title:fresh",
        "p[style-name='Subtitle'] => h2.subtitle:fresh",
        "r[style-name='Strong'] => strong",
        "r[style-name='Emphasis'] => em",
    ],
    // 图片转换选项
    convertImage: mammoth.images.imgElement(function(image) {
        return image.read('base64').then(function(imageBuffer) {
            return {
                src: "data:" + image.contentType + ";base64," + imageBuffer
            };
        });
    }),
    // 忽略样式
    ignoreEmptyParagraphs: false,
    // 保留换行
    preserveParagraphAlignment: true,
    preserveStyles: true,
    // 保留字符样式
    preserveCharacterStyles: true,
};

// 主函数
async function convert(inputPath, outputPath) {
    try {
        // 验证输入文件
        if (!fs.existsSync(inputPath)) {
            console.error(JSON.stringify({
                success: false,
                error: '输入文件不存在: ' + inputPath
            }));
            process.exit(1);
        }

        // 读取文件
        const buffer = fs.readFileSync(inputPath);
        
        // 转换为HTML
        const result = await mammoth.convertToHtml(
            { buffer: buffer },
            options
        );

        // 提取消息
        const messages = result.messages
            .filter(m => m.type === 'message')
            .map(m => m.message);
        
        const warnings = result.messages
            .filter(m => m.type === 'warning')
            .map(m => m.message);

        // 构建输出
        const output = {
            success: true,
            html: result.value,
            messages: messages,
            warnings: warnings,
            stats: {
                paragraphs: (result.value.match(/<p/g) || []).length,
                images: (result.value.match(/<img/g) || []).length,
                tables: (result.value.match(/<table/g) || []).length
            }
        };

        // 输出结果
        if (outputPath) {
            fs.writeFileSync(outputPath, JSON.stringify(output, null, 2));
            console.log('转换完成，输出文件:', outputPath);
        } else {
            // 输出到stdout
            process.stdout.write(JSON.stringify(output));
        }

    } catch (error) {
        console.error(JSON.stringify({
            success: false,
            error: error.message,
            stack: error.stack
        }));
        process.exit(1);
    }
}

// 获取文档元信息
async function getMetadata(inputPath) {
    try {
        const buffer = fs.readFileSync(inputPath);
        const result = await mammoth.extractRawText({ buffer: buffer });
        
        const stats = {
            characters: result.value.length,
            words: result.value.split(/\s+/).filter(w => w.length > 0).length,
            paragraphs: (result.value.match(/\n\n+/g) || []).length + 1
        };

        console.log(JSON.stringify({
            success: true,
            stats: stats
        }));

    } catch (error) {
        console.error(JSON.stringify({
            success: false,
            error: error.message
        }));
        process.exit(1);
    }
}

// 命令行入口
const args = process.argv.slice(2);
if (args.length < 1) {
    console.error('用法: node docx2html.js <input.docx> [output.json]');
    console.error('      node docx2html.js --metadata <input.docx>');
    process.exit(1);
}

const command = args[0];
if (command === '--metadata') {
    getMetadata(args[1]);
} else {
    convert(args[0], args[1]);
}
