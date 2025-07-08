const { getLanguageType } = require('./fileHandler');

/**
 * 解析代码文件，移除注释和空行
 * @param {Array} files - 文件信息数组
 * @param {Object} options - 解析选项
 * @returns {Promise<Object>} 解析结果
 */
async function parseCode(files, options = {}) {
    const {
        removeComments = true,
        removeEmptyLines = true
    } = options;
    
    let totalContent = '';
    let totalLines = 0;
    let processedFiles = 0;
    
    for (const file of files) {
        try {
            const languageType = getLanguageType(file.path);
            let content = file.content;

            if (removeComments) {
                content = removeCodeComments(content, languageType);
            }

            if (removeEmptyLines) {
                content = removeEmptyLinesFromCode(content);
            }

            // 如果处理后还有内容，则添加到总内容中
            if (content.trim()) {
                // 文件之间紧密连接，不添加额外空行
                if (totalContent) {
                    // 确保上一个文件内容不以空行结尾，然后添加新文件
                    totalContent = totalContent.trimEnd() + '\n' + content.trim();
                    totalLines += content.split('\n').length;
                } else {
                    totalContent = content.trim();
                    totalLines += content.split('\n').length;
                }
            }

            processedFiles++;
        } catch (error) {
            console.error(`处理文件 ${file.path} 时出错:`, error);
        }
    }
    
    return {
        content: totalContent,
        totalLines: totalLines,
        processedFiles: processedFiles,
        originalFiles: files.length
    };
}

/**
 * 移除代码中的注释
 * @param {string} content - 代码内容
 * @param {string} languageType - 编程语言类型
 * @returns {string} 移除注释后的代码
 */
function removeCodeComments(content, languageType) {
    switch (languageType) {
        case 'javascript':
        case 'typescript':
        case 'java':
        case 'cpp':
        case 'c':
        case 'csharp':
        case 'go':
        case 'rust':
        case 'php':
            return removeCStyleComments(content);
        
        case 'python':
            return removePythonComments(content);
        
        case 'html':
        case 'xml':
            return removeHtmlComments(content);
        
        case 'css':
        case 'scss':
        case 'sass':
            return removeCssComments(content);
        
        case 'shell':
            return removeShellComments(content);
        
        case 'sql':
            return removeSqlComments(content);
        
        default:
            return content;
    }
}

/**
 * 移除C风格注释
 * @param {string} content - 代码内容
 * @returns {string} 移除注释后的代码
 */
function removeCStyleComments(content) {
    // 移除单行注释
    content = content.replace(/\/\/.*$/gm, '');

    // 移除多行注释
    content = content.replace(/\/\*[\s\S]*?\*\//g, '');

    return content;
}

/**
 * 移除Python注释
 * @param {string} content - 代码内容
 * @returns {string} 移除注释后的代码
 */
function removePythonComments(content) {
    const lines = content.split('\n');
    const result = [];
    let inMultilineString = false;
    let stringDelimiter = '';
    
    for (let line of lines) {
        let processedLine = '';
        let i = 0;
        
        while (i < line.length) {
            const char = line[i];
            const nextChar = line[i + 1];
            const next2Chars = line.substring(i, i + 3);
            
            if (!inMultilineString) {
                // 检查多行字符串开始
                if (next2Chars === '"""' || next2Chars === "'''") {
                    inMultilineString = true;
                    stringDelimiter = next2Chars;
                    processedLine += next2Chars;
                    i += 3;
                    continue;
                }
                
                // 检查单行注释
                if (char === '#') {
                    break; // 忽略行的其余部分
                }
                
                // 检查字符串
                if (char === '"' || char === "'") {
                    // 处理字符串，跳过其中的内容
                    const quote = char;
                    processedLine += char;
                    i++;
                    
                    while (i < line.length && line[i] !== quote) {
                        if (line[i] === '\\') {
                            processedLine += line[i];
                            i++;
                            if (i < line.length) {
                                processedLine += line[i];
                                i++;
                            }
                        } else {
                            processedLine += line[i];
                            i++;
                        }
                    }
                    
                    if (i < line.length) {
                        processedLine += line[i]; // 添加结束引号
                        i++;
                    }
                    continue;
                }
            } else {
                // 在多行字符串中，检查结束
                if (next2Chars === stringDelimiter) {
                    processedLine += next2Chars;
                    i += 3;
                    inMultilineString = false;
                    stringDelimiter = '';
                    continue;
                }
            }
            
            processedLine += char;
            i++;
        }
        
        result.push(processedLine);
    }
    
    return result.join('\n');
}

/**
 * 移除HTML/XML注释
 * @param {string} content - 代码内容
 * @returns {string} 移除注释后的代码
 */
function removeHtmlComments(content) {
    return content.replace(/<!--[\s\S]*?-->/g, '');
}

/**
 * 移除CSS注释
 * @param {string} content - 代码内容
 * @returns {string} 移除注释后的代码
 */
function removeCssComments(content) {
    return content.replace(/\/\*[\s\S]*?\*\//g, '');
}

/**
 * 移除Shell注释
 * @param {string} content - 代码内容
 * @returns {string} 移除注释后的代码
 */
function removeShellComments(content) {
    return content.replace(/^#.*$/gm, '');
}

/**
 * 移除SQL注释
 * @param {string} content - 代码内容
 * @returns {string} 移除注释后的代码
 */
function removeSqlComments(content) {
    // 移除单行注释
    content = content.replace(/--.*$/gm, '');

    // 移除多行注释
    content = content.replace(/\/\*[\s\S]*?\*\//g, '');

    return content;
}

/**
 * 移除空行和多余空白
 * @param {string} content - 代码内容
 * @returns {string} 移除空行后的代码
 */
function removeEmptyLinesFromCode(content) {
    return content
        .split('\n')
        .filter(line => {
            const trimmed = line.trim();
            // 过滤掉完全空白的行和只有空白字符的行
            return trimmed !== '' && trimmed !== ' ' && trimmed !== '\t';
        })
        .map(line => line.trimEnd()) // 移除行尾空白
        .join('\n')
        .replace(/\n{3,}/g, '\n\n'); // 将连续3个或更多换行符替换为最多2个
}

/**
 * 统计代码行数
 * @param {string} content - 代码内容
 * @returns {number} 行数
 */
function countLines(content) {
    return content.split('\n').length;
}

module.exports = {
    parseCode,
    removeCodeComments,
    removeEmptyLinesFromCode,
    countLines
};
