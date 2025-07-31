const fs = require('fs-extra');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Header, Footer, PageNumber, AlignmentType } = require('docx');

/**
 * 根据页码模式处理内容
 * @param {string} content - 原始内容
 * @param {number} linesPerPage - 每页行数
 * @param {string} pageMode - 页码模式
 * @param {string} pageRange - 自定义页码范围
 * @returns {Object} 处理结果包含内容和页码信息
 */
function processContentByPageMode(content, linesPerPage, pageMode, pageRange = '') {
    const lines = content.split('\n');
    const totalPages = Math.ceil(lines.length / linesPerPage);
    
    if (pageMode === 'custom' && pageRange) {
        try {
            const customPages = parsePageRange(pageRange, totalPages);
            const selectedContent = [];
            
            customPages.forEach(pageNum => {
                const startLine = (pageNum - 1) * linesPerPage;
                const endLine = Math.min(startLine + linesPerPage, lines.length);
                const pageLines = lines.slice(startLine, endLine);
                selectedContent.push(...pageLines);
            });
            
            return {
                content: selectedContent.join('\n'),
                totalPages: totalPages,
                outputPages: customPages.length,
                customPages: customPages,
                pageRanges: customPages.map(pageNum => ({
                    page: pageNum,
                    start: (pageNum - 1) * linesPerPage,
                    end: Math.min(pageNum * linesPerPage, lines.length)
                }))
            };
        } catch (error) {
            throw new Error(`页码范围解析错误: ${error.message}`);
        }
    }
    
    // 全部页面
    return {
        content: content,
        totalPages: totalPages,
        outputPages: totalPages,
        pageRanges: [{ start: 1, end: totalPages, content: lines }]
    };
}

/**
 * 解析页码范围字符串
 * @param {string} rangeStr - 页码范围字符串
 * @param {number} totalPages - 总页数
 * @returns {Array} 页码数组
 */
function parsePageRange(rangeStr, totalPages) {
    const pages = [];
    const ranges = rangeStr.split(',');
    
    for (const range of ranges) {
        const trimmedRange = range.trim();
        
        if (trimmedRange.includes('-')) {
            const [start, end] = trimmedRange.split('-').map(num => parseInt(num.trim()));
            
            if (isNaN(start) || isNaN(end) || start < 1 || end > totalPages || start > end) {
                throw new Error(`无效的页码范围: ${trimmedRange}`);
            }
            
            for (let i = start; i <= end; i++) {
                if (!pages.includes(i)) {
                    pages.push(i);
                }
            }
        } else {
            const page = parseInt(trimmedRange);
            
            if (isNaN(page) || page < 1 || page > totalPages) {
                throw new Error(`无效的页码: ${trimmedRange}`);
            }
            
            if (!pages.includes(page)) {
                pages.push(page);
            }
        }
    }
    
    return pages.sort((a, b) => a - b);
}

/**
 * 生成文档
 * @param {Object} options - 生成选项
 * @returns {Promise<void>}
 */
async function generateDocument(options) {
    const {
        content,
        savePath,
        linesPerPage = 50,
        headerText = '',
        pageMode = 'all',
        pageRange = ''
    } = options;

    // 处理页码模式
    const processedResult = processContentByPageMode(content, linesPerPage, pageMode, pageRange);

    // 生成Word文档
    await generateWord(processedResult.content, savePath, {
        linesPerPage,
        headerText,
        pageMode,
        pageRange,
        pageInfo: processedResult
    });
}

/**
 * 生成Word文档
 * @param {string} content - 代码内容
 * @param {string} savePath - 保存路径
 * @param {Object} options - 选项
 */
async function generateWord(content, savePath, options) {
    const { linesPerPage, headerText, pageMode, pageInfo } = options;
    
    // 分割内容为行
    const lines = content.split('\n');
    const totalPages = pageInfo ? pageInfo.totalPages : Math.ceil(lines.length / linesPerPage);
    
    // 计算行间距，确保每页正好显示指定行数并铺满页面
    const availableHeight = 230 * 56.7; // A4可用高度(twips)
    const lineSpacing = Math.floor(availableHeight / linesPerPage);
    
    // 根据行间距调整字体大小，增大字号范围
    const fontSize = Math.max(18, Math.min(30, Math.floor(lineSpacing / 50))); // 字体大小在9-15pt之间 (18-30 half-points)
    
    // 创建段落数组
    const paragraphs = [];
    let currentLineInPage = 0;
    let currentOutputPage = 1;
    
    // 计算实际页码
    let actualPageNumber = 1;
    if (pageInfo && pageInfo.pageRanges) {
        actualPageNumber = pageInfo.pageRanges[0].start;
    }
    
    // 连续处理所有行，跳过空白行
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        
        // 跳过完全空白的行，避免产生不必要的空白
        if (!line || line.trim() === '') {
            continue;
        }
        
        // 检查是否需要分页
        if (currentLineInPage >= linesPerPage && i < lines.length - 1) {
            // 添加分页符，但不增加空白内容
            paragraphs.push(
                new Paragraph({
                    children: [new TextRun({ text: '' })],
                    pageBreakBefore: true,
                    spacing: { before: 0, after: 0 } // 移除分页前后的空白
                })
            );
            currentLineInPage = 0;
            currentOutputPage++;
            
            // 计算下一页的实际页码
            actualPageNumber = calculateActualPageNumber(currentOutputPage, pageInfo);
        }
        
        // 添加当前行，确保有实际内容
        paragraphs.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: line,
                        font: 'Times New Roman',
                        size: fontSize
                    })
                ],
                spacing: {
                    line: lineSpacing,
                    lineRule: 'exact', // 使用精确行距
                    before: 0, // 移除段落前空白
                    after: 0   // 移除段落后空白
                }
            })
        );
        
        currentLineInPage++;
    }
    
    // 创建页眉页脚
    const headerFooterOptions = createHeaderFooter(headerText, pageInfo);
    
    // 创建文档，使用紧凑的页面设置
    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    margin: {
                        top: 1080,    // 减小上边距 (0.75 inch)
                        right: 1080,  // 减小右边距 (0.75 inch)
                        bottom: 1080, // 减小下边距 (0.75 inch)
                        left: 1080    // 减小左边距 (0.75 inch)
                    }
                }
            },
            headers: headerFooterOptions.headers,
            footers: headerFooterOptions.footers,
            children: paragraphs
        }]
    });
    
    // 生成并保存文档
    const buffer = await Packer.toBuffer(doc);
    await fs.writeFile(savePath, buffer);
}

/**
 * 计算实际页码
 * @param {number} outputPage - 输出页码（1-60）
 * @param {Object} pageInfo - 页码信息
 * @returns {number} 实际页码
 */
function calculateActualPageNumber(outputPage, pageInfo) {
    if (!pageInfo || !pageInfo.pageRanges) {
        return outputPage;
    }
    
    if (pageInfo.pageRanges.length === 1) {
        // 全部页面模式
        return outputPage;
    } else {
        // 前30页+后30页模式
        if (outputPage <= 30) {
            // 前30页：1-30
            return outputPage;
        } else {
            // 后30页：从第(总页数-29)页开始
            const backStartPage = pageInfo.pageRanges[1].start;
            return backStartPage + (outputPage - 31);
        }
    }
}

/**
 * 创建页眉页脚
 * @param {string} headerText - 页眉文本
 * @param {Object} pageInfo - 页码信息
 * @returns {Object} 页眉页脚配置
 */
function createHeaderFooter(headerText, pageInfo) {
    const totalPages = pageInfo ? pageInfo.totalPages : 1;

    // 注意：Word文档的页码会在转换为PDF时保持正确的格式
    // 对于"前30页+后30页"模式，我们需要在Word中使用特殊的页码处理

    return {
        headers: {
            default: new Header({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: headerText,
                                font: 'SimSun',
                                size: 28, // 增大页眉字体大小到14pt
                                bold: true
                            })
                        ],
                        alignment: AlignmentType.CENTER
                    })
                ]
            })
        },
        footers: {
            default: new Footer({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                children: [PageNumber.CURRENT]
                            })
                        ],
                        alignment: AlignmentType.CENTER
                    })
                ]
            })
        }
    };
}

module.exports = {
    generateDocument
};
