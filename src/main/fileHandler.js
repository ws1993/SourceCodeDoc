const fs = require('fs-extra');
const path = require('path');

// 支持的代码文件扩展名
const SUPPORTED_EXTENSIONS = [
    '.js', '.jsx', '.ts', '.tsx',
    '.py', '.pyw',
    '.java',
    '.cpp', '.cxx', '.cc', '.c',
    '.cs',
    '.php',
    '.go',
    '.rs',
    '.vue',
    '.html', '.htm',
    '.css', '.scss', '.sass',
    '.json',
    '.xml',
    '.sql',
    '.sh', '.bat',
    '.md'
];

/**
 * 扫描指定路径中的所有代码文件
 * @param {string|string[]} paths - 文件或文件夹路径数组
 * @returns {Promise<Array>} 文件信息数组
 */
async function scanFiles(paths) {
    const files = [];

    // 确保paths是数组
    const pathArray = Array.isArray(paths) ? paths : [paths];

    for (const inputPath of pathArray) {
        try {
            const stat = await fs.stat(inputPath);
            
            if (stat.isFile()) {
                // 单个文件
                if (isCodeFile(inputPath)) {
                    const fileInfo = await getFileInfo(inputPath);
                    if (fileInfo) {
                        files.push(fileInfo);
                    }
                }
            } else if (stat.isDirectory()) {
                // 文件夹，递归扫描
                const dirFiles = await scanDirectory(inputPath);
                files.push(...dirFiles);
            }
        } catch (error) {
            console.error(`扫描路径 ${inputPath} 时出错:`, error);
        }
    }
    
    // 按文件路径排序
    files.sort((a, b) => a.path.localeCompare(b.path));
    
    return files;
}

/**
 * 递归扫描目录
 * @param {string} dirPath - 目录路径
 * @returns {Promise<Array>} 文件信息数组
 */
async function scanDirectory(dirPath) {
    const files = [];
    
    try {
        const entries = await fs.readdir(dirPath, { withFileTypes: true });
        
        for (const entry of entries) {
            const fullPath = path.join(dirPath, entry.name);
            
            // 跳过隐藏文件和特定目录
            if (shouldSkip(entry.name)) {
                continue;
            }
            
            if (entry.isFile()) {
                if (isCodeFile(fullPath)) {
                    const fileInfo = await getFileInfo(fullPath);
                    if (fileInfo) {
                        files.push(fileInfo);
                    }
                }
            } else if (entry.isDirectory()) {
                // 递归扫描子目录
                const subFiles = await scanDirectory(fullPath);
                files.push(...subFiles);
            }
        }
    } catch (error) {
        console.error(`扫描目录 ${dirPath} 时出错:`, error);
    }
    
    return files;
}

/**
 * 获取文件信息
 * @param {string} filePath - 文件路径
 * @returns {Promise<Object|null>} 文件信息对象
 */
async function getFileInfo(filePath) {
    try {
        const stat = await fs.stat(filePath);
        const content = await fs.readFile(filePath, 'utf8');
        const lines = content.split('\n').length;
        
        return {
            path: filePath,
            name: path.basename(filePath),
            extension: path.extname(filePath),
            size: stat.size,
            lines: lines,
            content: content,
            lastModified: stat.mtime
        };
    } catch (error) {
        console.error(`读取文件 ${filePath} 时出错:`, error);
        return null;
    }
}

/**
 * 判断是否为代码文件
 * @param {string} filePath - 文件路径
 * @returns {boolean} 是否为代码文件
 */
function isCodeFile(filePath) {
    const ext = path.extname(filePath).toLowerCase();
    return SUPPORTED_EXTENSIONS.includes(ext);
}

/**
 * 判断是否应该跳过文件或目录
 * @param {string} name - 文件或目录名
 * @returns {boolean} 是否跳过
 */
function shouldSkip(name) {
    // 跳过的目录和文件模式
    const skipPatterns = [
        // 隐藏文件
        /^\./,
        // 依赖目录
        /^node_modules$/,
        /^vendor$/,
        /^__pycache__$/,
        /^\.git$/,
        /^\.svn$/,
        /^\.hg$/,
        // 构建目录
        /^build$/,
        /^dist$/,
        /^target$/,
        /^bin$/,
        /^obj$/,
        // 临时文件
        /~$/,
        /\.tmp$/,
        /\.temp$/,
        // 日志文件
        /\.log$/,
        // 其他
        /^coverage$/,
        /^\.nyc_output$/
    ];
    
    return skipPatterns.some(pattern => pattern.test(name));
}

/**
 * 获取文件的编程语言类型
 * @param {string} filePath - 文件路径
 * @returns {string} 编程语言类型
 */
function getLanguageType(filePath) {
    const ext = path.extname(filePath).toLowerCase();
    
    const languageMap = {
        '.js': 'javascript',
        '.jsx': 'javascript',
        '.ts': 'typescript',
        '.tsx': 'typescript',
        '.py': 'python',
        '.pyw': 'python',
        '.java': 'java',
        '.cpp': 'cpp',
        '.cxx': 'cpp',
        '.cc': 'cpp',
        '.c': 'c',
        '.cs': 'csharp',
        '.php': 'php',
        '.go': 'go',
        '.rs': 'rust',
        '.vue': 'vue',
        '.html': 'html',
        '.htm': 'html',
        '.css': 'css',
        '.scss': 'scss',
        '.sass': 'sass',
        '.json': 'json',
        '.xml': 'xml',
        '.sql': 'sql',
        '.sh': 'shell',
        '.bat': 'batch',
        '.md': 'markdown'
    };
    
    return languageMap[ext] || 'text';
}

module.exports = {
    scanFiles,
    getLanguageType,
    isCodeFile,
    SUPPORTED_EXTENSIONS
};
