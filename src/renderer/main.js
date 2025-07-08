// 应用状态
let selectedFiles = [];
let selectedPaths = new Set(); // 用于跟踪已选择的路径，避免重复
let parsedCode = '';
let currentSettings = {
    linesPerPage: 50,
    headerText: '源程序V1.0',
    outputFormat: 'docx',
    pageMode: 'all',
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

// 检查必需的DOM元素是否存在
function checkRequiredElements() {
    const requiredElements = [
        'selectFiles', 'selectFolder', 'clearSelection', 'dropZone',
        'pathList', 'pathItems', 'fileList', 'fileItems', 'fileCount', 'totalLines',
        'linesPerPage', 'headerText', 'pageMode', 'removeComments', 'removeEmptyLines',
        'preview', 'previewContent', 'estimatedPages', 'previewLines', 'generateBtn',
        'progressModal', 'progressText', 'statusText'
    ];

    const missingElements = [];
    for (const elementName of requiredElements) {
        if (!elements[elementName]) {
            missingElements.push(elementName);
        }
    }

    if (missingElements.length > 0) {
        console.error('缺少必需的DOM元素:', missingElements);
        showError(`页面初始化失败，缺少元素: ${missingElements.join(', ')}`);
        return false;
    }

    return true;
}

// 初始化事件监听器
function initializeEventListeners() {
    // 检查必需元素
    if (!checkRequiredElements()) {
        return;
    }

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
    elements.pageMode.addEventListener('change', updateSettings);
    elements.removeComments.addEventListener('change', updateSettings);
    elements.removeEmptyLines.addEventListener('change', updateSettings);
    
    // 生成按钮
    elements.generateBtn.addEventListener('click', handleGenerate);
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
    if (elements.generateBtn) {
        elements.generateBtn.disabled = true;
    }
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

// 移除单个文件
function removeFile(filePathToRemove) {
    console.log('removeFile called with:', filePathToRemove);
    console.log('Current selectedFiles count:', selectedFiles.length);

    // 记录移除前的文件数量
    const beforeCount = selectedFiles.length;

    // 从selectedFiles数组中移除指定文件
    selectedFiles = selectedFiles.filter(file => {
        const shouldKeep = file.path !== filePathToRemove;
        if (!shouldKeep) {
            console.log('Removing file:', file.path);
        }
        return shouldKeep;
    });

    const afterCount = selectedFiles.length;
    console.log('Files removed:', beforeCount - afterCount);

    // 更新文件列表显示
    updateFileList();

    // 重新解析和预览
    if (selectedFiles.length > 0) {
        parseAndPreview();
    } else {
        // 如果没有文件了，隐藏预览
        elements.preview.style.display = 'none';
        if (elements.generateBtn) {
            elements.generateBtn.disabled = true;
        }
        parsedCode = '';
    }

    updateStatus(`已移除文件，剩余 ${selectedFiles.length} 个文件`);
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

    selectedFiles.forEach((file, index) => {
        const item = document.createElement('div');
        item.className = 'file-item';

        // 创建删除按钮，使用事件监听器而不是onclick属性
        const removeButton = document.createElement('button');
        removeButton.className = 'remove-file';
        removeButton.textContent = '删除';
        removeButton.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            console.log('Delete button clicked for file:', file.path);
            removeFile(file.path);
        });

        // 创建其他元素
        const fileIcon = document.createElement('span');
        fileIcon.className = 'file-icon';
        fileIcon.textContent = '📄';

        const fileName = document.createElement('span');
        fileName.className = 'file-name';
        fileName.textContent = file.name;

        const filePath = document.createElement('span');
        filePath.className = 'file-path';
        filePath.textContent = file.path;

        const fileLines = document.createElement('span');
        fileLines.className = 'file-lines';
        fileLines.textContent = `${file.lines || 0} 行`;

        // 组装元素
        item.appendChild(fileIcon);
        item.appendChild(fileName);
        item.appendChild(filePath);
        item.appendChild(fileLines);
        item.appendChild(removeButton);

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
        if (elements.generateBtn) {
            elements.generateBtn.disabled = true;
        }
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
        if (elements.generateBtn) {
            elements.generateBtn.disabled = false;
        }
        
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
    let pageRangeText = '';

    if (currentSettings.pageMode === 'partial' && totalPages > 60) {
        outputPages = 60; // 前30页 + 后30页
        const backStartPage = totalPages - 29; // 后30页起始页码
        pageRangeText = ` (1-30页 + ${backStartPage}-${totalPages}页，共${totalPages}页)`;
    } else {
        pageRangeText = ` (全部页面)`;
    }

    elements.previewLines.textContent = totalLines;
    elements.estimatedPages.textContent = `${outputPages}页输出${pageRangeText}`;

    // 显示前几行作为预览
    const previewLines = parseResult.content.split('\n').slice(0, 20);
    elements.previewContent.textContent = previewLines.join('\n') + '\n...(更多内容)';
}

// 更新设置
function updateSettings() {
    currentSettings.linesPerPage = parseInt(elements.linesPerPage.value);
    currentSettings.headerText = elements.headerText.value;
    currentSettings.pageMode = elements.pageMode.value;
    currentSettings.removeComments = elements.removeComments.checked;
    currentSettings.removeEmptyLines = elements.removeEmptyLines.checked;
    
    // 重新解析和预览
    if (selectedFiles.length > 0) {
        parseAndPreview();
    }
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
window.removeFile = removeFile;

// 初始化应用
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    updateStatus('就绪');
});
