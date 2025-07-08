const fs = require('fs-extra');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Header, Footer, PageNumber, AlignmentType } = require('docx');

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
        headerText = '源程序V1.0'
    } = options;

    // 直接生成Word文档
    await generateWord(content, savePath, {
        linesPerPage,
        headerText
    });
}





/**
 * 生成Word文档
 * @param {string} content - 代码内容
 * @param {string} savePath - 保存路径
 * @param {Object} options - 选项
 */
async function generateWord(content, savePath, options) {
    const { linesPerPage, headerText } = options;

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
    const paragraphResult = createParagraphsFromLines(lines, linesPerPage, fontSize, lineSpacing);
    const paragraphs = paragraphResult.paragraphs;
    const realActualPages = paragraphResult.actualPages;

    // 创建页眉页脚，使用实际生成的页数
    const headerFooterOptions = createHeaderFooter(headerText, {
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
 * 从行数组创建段落数组
 * @param {string[]} lines - 行数组
 * @param {number} linesPerPage - 每页行数
 * @param {number} fontSize - 字体大小
 * @param {number} lineSpacing - 行间距
 * @returns {Object} 包含段落数组和实际页数的对象
 */
function createParagraphsFromLines(lines, linesPerPage, fontSize, lineSpacing) {
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
        actualPages = Math.ceil(contentLineCount / linesPerPage);
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



module.exports = {
    generateDocument
};
