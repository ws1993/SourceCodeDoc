// 应用状态
let selectedFiles = [];
let selectedPaths = new Set(); // 用于跟踪已选择的路径，避免重复
let parsedCode = '';
let currentSettings = {
    linesPerPage: 50,
    headerText: '',
    outputFormat: 'docx',
    pageMode: 'all',
    pageRange: '',
    removeComments: true,
    removeEmptyLines: true
};

// DOM元素
const elements = {
    selectFiles: document.getElementById('selectFiles'),
    selectFolder: document.getElementById('selectFolder'),
    clearSelection: document.getElementById('clearSelection'),
    dropZone: document.getElementById('dropZone'),
    pathList: document.getElementById('pathList'),
    pathItems: document.getElementById('pathItems'),
    fileList: document.getElementById('fileList'),
    fileItems: document.getElementById('fileItems'),
    fileCount: document.getElementById('fileCount'),
    totalLines: document.getElementById('totalLines'),
    linesPerPage: document.getElementById('linesPerPage'),
    headerText: document.getElementById('headerText'),
    pageMode: document.getElementById('pageMode'),
    pageRange: document.getElementById('pageRange'),
    customPageRange: document.getElementById('customPageRange'),
    removeComments: document.getElementById('removeComments'),
    removeEmptyLines: document.getElementById('removeEmptyLines'),
    preview: document.querySelector('.preview'),
    previewContent: document.getElementById('previewContent'),
    estimatedPages: document.getElementById('estimatedPages'),
    previewLines: document.getElementById('previewLines'),
    generateBtn: document.getElementById('generateBtn'),
    progressModal: document.getElementById('progressModal'),
    progressText: document.getElementById('progressText'),
    statusText: document.getElementById('statusText')
};

// 初始化事件监听器
function initializeEventListeners() {
    // 文件选择按钮
    elements.selectFiles.addEventListener('click', handleSelectFiles);
    elements.selectFolder.addEventListener('click', handleSelectFolder);
    elements.clearSelection.addEventListener('click', handleClearSelection);
    
    // 拖拽功能
    elements.dropZone.addEventListener('dragover', handleDragOver);
    elements.dropZone.addEventListener('dragleave', handleDragLeave);
    elements.dropZone.addEventListener('drop', handleDrop);
    
    // 设置变更
    elements.linesPerPage.addEventListener('change', updateSettings);
    elements.headerText.addEventListener('input', updateSettings);
    elements.pageMode.addEventListener('change', handlePageModeChange);
    elements.pageRange.addEventListener('input', updateSettings);
    elements.removeComments.addEventListener('change', updateSettings);
    elements.removeEmptyLines.addEventListener('change', updateSettings);
    
    // 生成按钮
    elements.generateBtn.addEventListener('click', handleGenerate);
}

// 处理页码模式变更
function handlePageModeChange() {
    const pageMode = elements.pageMode.value;
    
    if (pageMode === 'custom') {
        elements.customPageRange.style.display = 'block';
    } else {
        elements.customPageRange.style.display = 'none';
    }
    
    updateSettings();
}

// 处理文件选择
async function handleSelectFiles() {
    try {
        const result = await window.electronAPI.selectFiles();
        if (!result.canceled && result.filePaths.length > 0) {
            await processSelectedPaths(result.filePaths);
        }
    } catch (error) {
        showError('选择文件时出错：' + error.message);
    }
}

// 处理文件夹选择
async function handleSelectFolder() {
    try {
        const result = await window.electronAPI.selectFolder();
        if (!result.canceled && result.filePaths.length > 0) {
            await processSelectedPaths(result.filePaths);
        }
    } catch (error) {
        showError('选择文件夹时出错：' + error.message);
    }
}

// 处理拖拽
function handleDragOver(e) {
    e.preventDefault();
    elements.dropZone.classList.add('drag-over');
}

function handleDragLeave(e) {
    e.preventDefault();
    elements.dropZone.classList.remove('drag-over');
}

async function handleDrop(e) {
    e.preventDefault();
    elements.dropZone.classList.remove('drag-over');
    
    const files = Array.from(e.dataTransfer.files);
    const paths = files.map(file => file.path);
    
    if (paths.length > 0) {
        await processSelectedPaths(paths);
    }
}

// 处理选择的路径
async function processSelectedPaths(paths, isRescan = false) {
    showProgress('正在扫描文件...');

    try {
        let pathsToScan = paths;

        if (!isRescan) {
            // 过滤掉重复的路径
            const newPaths = paths.filter(path => {
                const normalizedPath = path.replace(/\\/g, '/').toLowerCase();
                return !selectedPaths.has(normalizedPath);
            });

            if (newPaths.length === 0) {
                hideProgress();
                showError('所选择的文件夹已经添加过了');
                return;
            }

            pathsToScan = newPaths;

            // 添加新路径到已选择集合
            newPaths.forEach(path => {
                const normalizedPath = path.replace(/\\/g, '/').toLowerCase();
                selectedPaths.add(normalizedPath);
            });
        } else {
            // 重新扫描模式：清空现有文件列表
            selectedFiles = [];
            selectedPaths.clear();

            // 重新添加路径
            pathsToScan.forEach(path => {
                selectedPaths.add(path);
            });
        }

        // 扫描路径中的文件
        const scannedFiles = await window.electronAPI.scanFiles(pathsToScan);

        if (!isRescan) {
            // 合并新文件到现有文件列表
            selectedFiles = [...selectedFiles, ...scannedFiles];
        } else {
            // 重新扫描模式：直接使用新扫描的文件
            selectedFiles = scannedFiles;
        }

        // 去重（基于文件路径）
        const uniqueFiles = [];
        const seenPaths = new Set();

        selectedFiles.forEach(file => {
            const normalizedPath = file.path.replace(/\\/g, '/').toLowerCase();
            if (!seenPaths.has(normalizedPath)) {
                seenPaths.add(normalizedPath);
                uniqueFiles.push(file);
            }
        });

        selectedFiles = uniqueFiles;

        updatePathList();
        updateFileList();
        await parseAndPreview();

        hideProgress();
        updateStatus(`已选择 ${selectedFiles.length} 个文件，来自 ${selectedPaths.size} 个位置`);
    } catch (error) {
        hideProgress();
        showError('扫描文件时出错：' + error.message);
    }
}

// 清空选择
function handleClearSelection() {
    selectedFiles = [];
    selectedPaths.clear();
    parsedCode = '';

    updatePathList();
    updateFileList();
    elements.preview.style.display = 'none';
    elements.generateBtn.disabled = true;
    elements.clearSelection.style.display = 'none';

    updateStatus('已清空所有选择');
}

// 移除单个路径
function removePath(pathToRemove) {
    const normalizedPathToRemove = pathToRemove.replace(/\\/g, '/').toLowerCase();
    selectedPaths.delete(normalizedPathToRemove);

    // 重新扫描剩余路径的文件
    if (selectedPaths.size > 0) {
        const remainingPaths = Array.from(selectedPaths);
        processSelectedPaths(remainingPaths, true); // true表示重新扫描
    } else {
        handleClearSelection();
    }
}

// 更新路径列表显示
function updatePathList() {
    if (selectedPaths.size === 0) {
        elements.pathList.style.display = 'none';
        return;
    }

    elements.pathList.style.display = 'block';
    elements.pathItems.innerHTML = '';

    selectedPaths.forEach(path => {
        const item = document.createElement('div');
        item.className = 'path-item';
        item.innerHTML = `
            <span class="path-text" title="${path}">${path}</span>
            <button class="remove-path" onclick="removePath('${path.replace(/'/g, "\\'")}')">移除</button>
        `;
        elements.pathItems.appendChild(item);
    });
}

// 更新文件列表显示
function updateFileList() {
    if (selectedFiles.length === 0) {
        elements.fileList.style.display = 'none';
        elements.clearSelection.style.display = 'none';
        return;
    }

    elements.fileList.style.display = 'block';
    elements.clearSelection.style.display = 'inline-flex';
    elements.fileItems.innerHTML = '';

    let totalLines = 0;

    selectedFiles.forEach(file => {
        const item = document.createElement('div');
        item.className = 'file-item';
        item.innerHTML = `
            <span class="file-icon">📄</span>
            <span class="file-name">${file.name}</span>
            <span class="file-path">${file.path}</span>
            <span class="file-lines">${file.lines || 0} 行</span>
        `;
        elements.fileItems.appendChild(item);
        totalLines += file.lines || 0;
    });

    elements.fileCount.textContent = selectedFiles.length;
    elements.totalLines.textContent = totalLines;
}

// 解析代码并预览
async function parseAndPreview() {
    if (selectedFiles.length === 0) {
        elements.preview.style.display = 'none';
        elements.generateBtn.disabled = true;
        return;
    }
    
    showProgress('正在解析代码...');
    
    try {
        const parseOptions = {
            removeComments: currentSettings.removeComments,
            removeEmptyLines: currentSettings.removeEmptyLines
        };
        
        const result = await window.electronAPI.parseCode(selectedFiles, parseOptions);
        parsedCode = result.content;
        
        updatePreview(result);
        elements.generateBtn.disabled = false;
        
        hideProgress();
    } catch (error) {
        hideProgress();
        showError('解析代码时出错：' + error.message);
    }
}

// 更新预览
function updatePreview(parseResult) {
    elements.preview.style.display = 'block';

    const totalLines = parseResult.totalLines;
    const totalPages = Math.ceil(totalLines / currentSettings.linesPerPage);

    // 根据页码模式计算实际输出页数和页码范围
    let outputPages = totalPages;
    let outputLines = totalLines;
    let pageRangeText = '';

    if (currentSettings.pageMode === 'custom' && currentSettings.pageRange) {
        try {
            const customPages = parsePageRange(currentSettings.pageRange, totalPages);
            outputPages = customPages.length;
            
            // 计算自定义页码范围对应的实际行数
            outputLines = calculateCustomPageLines(customPages, totalLines, currentSettings.linesPerPage);
            
            pageRangeText = ` (自定义页码：${currentSettings.pageRange}，输出${outputPages}页)`;
        } catch (error) {
            pageRangeText = ` (页码范围格式错误)`;
            outputPages = 0;
            outputLines = 0;
        }
    } else {
        pageRangeText = ` (全部页面)`;
    }

    elements.previewLines.textContent = outputLines;
    elements.estimatedPages.textContent = `${outputPages}页输出${pageRangeText}`;

    // 显示前几行作为预览
    const previewLines = parseResult.content.split('\n').slice(0, 20);
    elements.previewContent.textContent = previewLines.join('\n') + '\n...(更多内容)';
}

/**
 * 计算自定义页码范围对应的实际行数
 * @param {Array} customPages - 自定义页码数组
 * @param {number} totalLines - 总行数
 * @param {number} linesPerPage - 每页行数
 * @returns {number} 实际行数
 */
function calculateCustomPageLines(customPages, totalLines, linesPerPage) {
    let actualLines = 0;
    
    customPages.forEach(pageNum => {
        const startLine = (pageNum - 1) * linesPerPage;
        const endLine = Math.min(startLine + linesPerPage, totalLines);
        const pageLines = endLine - startLine;
        actualLines += pageLines;
    });
    
    return actualLines;
}

// 更新设置
function updateSettings() {
    currentSettings.linesPerPage = parseInt(elements.linesPerPage.value);
    currentSettings.headerText = elements.headerText.value;
    currentSettings.pageMode = elements.pageMode.value;
    currentSettings.pageRange = elements.pageRange.value;
    currentSettings.removeComments = elements.removeComments.checked;
    currentSettings.removeEmptyLines = elements.removeEmptyLines.checked;
    
    // 重新解析和预览
    if (selectedFiles.length > 0) {
        parseAndPreview();
    }
}

// 解析页码范围字符串
function parsePageRange(rangeStr, totalPages) {
    const pages = [];
    const ranges = rangeStr.split(',');
    
    for (const range of ranges) {
        const trimmedRange = range.trim();
        
        if (trimmedRange.includes('-')) {
            // 处理范围，如 "1-20"
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
            // 处理单个页码，如 "5"
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

// 生成文档
async function handleGenerate() {
    if (!parsedCode) {
        showError('请先选择文件');
        return;
    }
    
    try {
        // 选择保存位置 - 只支持Word格式
        const saveResult = await window.electronAPI.saveDocument({
            format: 'docx'
        });

        if (saveResult.canceled) {
            return;
        }

        showProgress('正在生成Word文档...');

        const generateOptions = {
            ...currentSettings,
            savePath: saveResult.filePath,
            content: parsedCode
        };

        await window.electronAPI.generateDocument(generateOptions);

        hideProgress();
        showSuccess('Word文档生成成功！');
        updateStatus('Word文档生成完成');

    } catch (error) {
        hideProgress();
        showError('生成文档时出错：' + error.message);
    }
}

// 显示进度
function showProgress(text) {
    elements.progressText.textContent = text;
    elements.progressModal.style.display = 'flex';
}

// 隐藏进度
function hideProgress() {
    elements.progressModal.style.display = 'none';
}

// 更新状态
function updateStatus(text) {
    elements.statusText.textContent = text;
}

// 显示错误
function showError(message) {
    alert('错误：' + message);
    updateStatus('错误：' + message);
}

// 显示成功
function showSuccess(message) {
    alert(message);
    updateStatus(message);
}

// 将函数暴露到全局作用域
window.removePath = removePath;

// 初始化应用
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    updateStatus('就绪');
});
