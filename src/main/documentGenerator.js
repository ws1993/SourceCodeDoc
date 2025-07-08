const fs = require('fs-extra');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Header, Footer, PageNumber, AlignmentType, PageNumberStart } = require('docx');

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
        headerText = '源程序V1.0',
        pageMode = 'all'
    } = options;

    // 处理页码模式
    const processedResult = processContentByPageMode(content, linesPerPage, pageMode);

    // 只生成Word文档
    await generateWord(processedResult.content, savePath, {
        linesPerPage,
        headerText,
        pageMode,
        pageInfo: processedResult
    });
}

/**
 * 根据页码模式处理内容
 * @param {string} content - 原始内容
 * @param {number} linesPerPage - 每页行数
 * @param {string} pageMode - 页码模式
 * @returns {Object} 处理结果包含内容和页码信息
 */
function processContentByPageMode(content, linesPerPage, pageMode) {
    const lines = content.split('\n');

    // 先计算完整文档的页面分布
    const pageDistribution = calculatePageDistribution(lines, linesPerPage);
    const totalPages = pageDistribution.totalPages;

    if (pageMode === 'partial' && totalPages > 60) {
        // 前30页 + 后30页模式
        // 提取前30页的内容
        const firstPageContent = extractPagesContent(pageDistribution, 1, 30);

        // 计算后30页的起始页码
        const backStartPage = totalPages - 29; // 后30页从第(总页数-29)页开始

        // 提取后30页的内容
        const lastPageContent = extractPagesContent(pageDistribution, backStartPage, totalPages);

        return {
            content: [...firstPageContent, ...lastPageContent].join('\n'),
            totalPages: totalPages,
            outputPages: 60,
            pageRanges: [
                { start: 1, end: 30, content: firstPageContent },
                { start: backStartPage, end: totalPages, content: lastPageContent }
            ]
        };
    }

    // 全部页面或总页数不超过60页
    return {
        content: content,
        totalPages: totalPages,
        outputPages: totalPages,
        pageRanges: [{ start: 1, end: totalPages, content: lines }]
    };
}

/**
 * 计算页面分布
 * @param {string[]} lines - 所有行
 * @param {number} linesPerPage - 每页行数
 * @returns {Object} 页面分布信息
 */
function calculatePageDistribution(lines, linesPerPage) {
    const pages = [];
    let currentPage = 1;
    let currentLineInPage = 0;
    let pageStartIndex = 0;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        // 跳过完全空白的行，但记录它们的位置
        if (!line || line.trim() === '') {
            continue;
        }

        // 检查是否需要分页
        if (currentLineInPage >= linesPerPage && currentLineInPage > 0) {
            // 保存当前页的信息
            pages.push({
                pageNumber: currentPage,
                startIndex: pageStartIndex,
                endIndex: i - 1,
                lines: lines.slice(pageStartIndex, i)
            });

            // 开始新页
            currentPage++;
            currentLineInPage = 0;
            pageStartIndex = i;
        }

        currentLineInPage++;
    }

    // 添加最后一页
    if (pageStartIndex < lines.length) {
        pages.push({
            pageNumber: currentPage,
            startIndex: pageStartIndex,
            endIndex: lines.length - 1,
            lines: lines.slice(pageStartIndex)
        });
    }

    return {
        totalPages: currentPage,
        pages: pages,
        allLines: lines
    };
}

/**
 * 提取指定页面范围的内容
 * @param {Object} pageDistribution - 页面分布信息
 * @param {number} startPage - 起始页码
 * @param {number} endPage - 结束页码
 * @returns {string[]} 提取的行内容
 */
function extractPagesContent(pageDistribution, startPage, endPage) {
    const extractedLines = [];

    for (const page of pageDistribution.pages) {
        if (page.pageNumber >= startPage && page.pageNumber <= endPage) {
            extractedLines.push(...page.lines);
        }
    }

    return extractedLines;
}

/**
 * 生成Word文档
 * @param {string} content - 代码内容
 * @param {string} savePath - 保存路径
 * @param {Object} options - 选项
 */
async function generateWord(content, savePath, options) {
    const { linesPerPage, headerText, pageMode, pageInfo } = options;

    if (pageMode === 'partial' && pageInfo && pageInfo.pageRanges && pageInfo.pageRanges.length === 2) {
        // 前30页+后30页模式，使用多section结构
        await generateWordWithMultipleSections(content, savePath, options);
    } else {
        // 全部页面模式，使用单section结构
        await generateWordWithSingleSection(content, savePath, options);
    }
}

/**
 * 生成单section的Word文档（全部页面模式）
 * @param {string} content - 代码内容
 * @param {string} savePath - 保存路径
 * @param {Object} options - 选项
 */
async function generateWordWithSingleSection(content, savePath, options) {
    const { linesPerPage, headerText, pageInfo } = options;

    // 分割内容为行
    const lines = content.split('\n');

    // 计算实际页数（基于非空行）
    const nonEmptyLines = lines.filter(line => line && line.trim() !== '');
    const actualTotalPages = Math.ceil(nonEmptyLines.length / linesPerPage);

    // 计算行间距，确保每页正好显示指定行数并铺满页面
    const availableHeight = 230 * 56.7; // A4可用高度(twips)
    const lineSpacing = Math.floor(availableHeight / linesPerPage);

    // 根据行间距调整字体大小
    const fontSize = Math.max(18, Math.min(30, Math.floor(lineSpacing / 50)));

    // 创建段落数组
    const paragraphResult = createParagraphsFromLines(lines, linesPerPage, fontSize, lineSpacing, pageInfo);
    const paragraphs = paragraphResult.paragraphs;
    const realActualPages = paragraphResult.actualPages;

    // 创建页眉页脚，使用实际生成的页数
    const headerFooterOptions = createHeaderFooter(headerText, {
        ...pageInfo,
        totalPages: realActualPages,
        outputPages: realActualPages
    });

    // 创建文档
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
 * 生成多section的Word文档（前30页+后30页模式）
 * @param {string} content - 代码内容
 * @param {string} savePath - 保存路径
 * @param {Object} options - 选项
 */
async function generateWordWithMultipleSections(content, savePath, options) {
    const { linesPerPage, headerText, pageInfo } = options;

    // 计算行间距和字体大小
    const availableHeight = 230 * 56.7;
    const lineSpacing = Math.floor(availableHeight / linesPerPage);
    const fontSize = Math.max(18, Math.min(30, Math.floor(lineSpacing / 50)));

    const sections = [];

    // 处理每个页面范围
    for (let rangeIndex = 0; rangeIndex < pageInfo.pageRanges.length; rangeIndex++) {
        const range = pageInfo.pageRanges[rangeIndex];
        const rangeContent = Array.isArray(range.content) ? range.content.join('\n') : range.content;
        const rangeLines = rangeContent.split('\n');

        // 创建该范围的段落
        const paragraphResult = createParagraphsFromLines(rangeLines, linesPerPage, fontSize, lineSpacing, pageInfo);
        const paragraphs = paragraphResult.paragraphs;

        // 创建该范围的页眉页脚
        const headerFooterOptions = createHeaderFooterForRange(headerText, pageInfo, range, rangeIndex);

        // 创建section属性，设置页码起始值
        const sectionProperties = {
            page: {
                margin: {
                    top: 1080,
                    right: 1080,
                    bottom: 1080,
                    left: 1080
                }
            },
            // 为后30页section设置页码起始值
            pageNumberStart: rangeIndex === 1 ? range.start : undefined
        };

        // 创建section
        const section = {
            properties: sectionProperties,
            headers: headerFooterOptions.headers,
            footers: headerFooterOptions.footers,
            children: paragraphs
        };

        sections.push(section);
    }

    // 创建文档
    const doc = new Document({ sections });

    // 生成并保存文档
    const buffer = await Packer.toBuffer(doc);
    await fs.writeFile(savePath, buffer);
}

/**
 * 从行数组创建段落数组
 * @param {string[]} lines - 行数组
 * @param {number} linesPerPage - 每页行数
 * @param {number} fontSize - 字体大小
 * @param {number} lineSpacing - 行间距
 * @returns {Object} 包含段落数组和实际页数的对象
 */
function createParagraphsFromLines(lines, linesPerPage, fontSize, lineSpacing, pageInfo = null) {
    const paragraphs = [];
    let currentLineInPage = 0;
    let actualPages = 0;
    let contentLineCount = 0; // 实际内容行数

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        // 跳过完全空白的行，避免产生不必要的空白
        if (!line || line.trim() === '') {
            continue;
        }

        contentLineCount++;

        // 检查是否需要分页
        if (currentLineInPage >= linesPerPage && contentLineCount > 1) {
            // 添加分页符到下一个段落
            currentLineInPage = 0;
        }

        // 添加当前行，确保有实际内容
        const paragraph = new Paragraph({
            children: [
                new TextRun({
                    text: line,
                    font: 'Times New Roman',
                    size: fontSize
                })
            ],
            spacing: {
                line: lineSpacing,
                lineRule: 'exact',
                before: 0,
                after: 0
            },
            // 如果是新页的第一行且不是第一行内容，添加分页符
            pageBreakBefore: currentLineInPage === 0 && contentLineCount > 1
        });

        paragraphs.push(paragraph);
        currentLineInPage++;
    }

    // 计算实际页数
    if (contentLineCount === 0) {
        actualPages = 0;
    } else {
        // 如果有页面信息，使用页面信息中的页数
        if (pageInfo && pageInfo.outputPages) {
            actualPages = pageInfo.outputPages;
        } else {
            actualPages = Math.ceil(contentLineCount / linesPerPage);
        }
    }

    return {
        paragraphs,
        actualPages
    };
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
 * 创建页眉页脚（单section模式）
 * @param {string} headerText - 页眉文本
 * @param {Object} pageInfo - 页码信息
 * @returns {Object} 页眉页脚配置
 */
function createHeaderFooter(headerText, pageInfo) {
    const totalPages = pageInfo ? pageInfo.totalPages : 1;

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
                            }),
                            new TextRun({
                                text: `/${totalPages}`
                            })
                        ],
                        alignment: AlignmentType.CENTER
                    })
                ]
            })
        }
    };
}

/**
 * 为特定范围创建页眉页脚（多section模式）
 * @param {string} headerText - 页眉文本
 * @param {Object} pageInfo - 页码信息
 * @param {Object} range - 当前范围信息
 * @param {number} rangeIndex - 范围索引
 * @returns {Object} 页眉页脚配置
 */
function createHeaderFooterForRange(headerText, pageInfo, range, rangeIndex) {
    const totalPages = pageInfo.totalPages;

    return {
        headers: {
            default: new Header({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: headerText,
                                font: 'SimSun',
                                size: 28,
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
                            }),
                            new TextRun({
                                text: `/${totalPages}`
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
