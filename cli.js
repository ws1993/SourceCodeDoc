#!/usr/bin/env node

const fs = require('fs-extra');
const path = require('path');
const { generateDocument } = require('./src/main/documentGenerator');

// 命令行参数处理
const args = process.argv.slice(2);

if (args.length === 0) {
    console.log('源程序文档生成软件');
    console.log('使用方法: cli.js <输入文件路径> [选项]');
    console.log('');
    console.log('选项:');
    console.log('  --output <路径>     输出文件路径 (默认: output.docx)');
    console.log('  --lines <数量>      每页行数 (默认: 50)');
    console.log('  --header <文本>     页眉文本 (默认: )');
    console.log('  --mode <模式>       页码模式: all|partial (默认: all)');
    console.log('');
    console.log('示例:');
    console.log('  cli.js input.txt --output output.docx --lines 45 --mode partial');
    process.exit(0);
}

// 解析参数
const inputFile = args[0];
let outputFile = 'output.docx';
let linesPerPage = 50;
let headerText = '';
let pageMode = 'all';

for (let i = 1; i < args.length; i += 2) {
    const option = args[i];
    const value = args[i + 1];
    
    switch (option) {
        case '--output':
            outputFile = value;
            break;
        case '--lines':
            linesPerPage = parseInt(value) || 50;
            break;
        case '--header':
            headerText = value;
            break;
        case '--mode':
            pageMode = value;
            break;
        default:
            console.log(`未知选项: ${option}`);
            process.exit(1);
    }
}

async function main() {
    try {
        // 检查输入文件是否存在
        if (!await fs.pathExists(inputFile)) {
            console.error(`错误: 输入文件不存在: ${inputFile}`);
            process.exit(1);
        }
        
        // 读取输入文件内容
        console.log(`正在读取文件: ${inputFile}`);
        const content = await fs.readFile(inputFile, 'utf-8');
        
        // 生成文档
        console.log(`正在生成文档...`);
        console.log(`- 输出文件: ${outputFile}`);
        console.log(`- 每页行数: ${linesPerPage}`);
        console.log(`- 页眉文本: ${headerText}`);
        console.log(`- 页码模式: ${pageMode}`);
        
        await generateDocument({
            content,
            savePath: outputFile,
            linesPerPage,
            headerText,
            pageMode
        });
        
        console.log(`✅ 文档生成成功: ${outputFile}`);
        
    } catch (error) {
        console.error(`❌ 生成文档时出错:`, error.message);
        process.exit(1);
    }
}

main();
