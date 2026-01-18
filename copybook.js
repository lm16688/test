// 字帖打印页面JavaScript

// 等待DOM完全加载
function waitForDOMReady() {
    return new Promise((resolve) => {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', resolve);
        } else {
            resolve();
        }
    });
}

// 主函数
async function main() {
    console.log('开始初始化字帖打印页面...');
    
    try {
        // 等待DOM完全加载
        await waitForDOMReady();
        console.log('DOM已加载');
        
        // 获取DOM元素 - 根据你的HTML实际ID
        const elements = {
            // 左侧选项区域
            gradeSelect: document.getElementById('grade-select'),
            generateBtn: document.getElementById('generate-btn'),
            characterCount: document.getElementById('character-count'),
            countNumber: document.getElementById('count-number'),
            uploadArea: document.getElementById('upload-area'),
            excelFileInput: document.getElementById('excel-file'),
            fileInfo: document.getElementById('file-info'),
            fileName: document.getElementById('file-name'),
            generateExcelBtn: document.getElementById('generate-excel-btn'),
            
            // 字帖预览区域
            copybookPreview: document.getElementById('copybook-preview'),
            previewTitle: document.getElementById('preview-title'),
            copybookGrid: document.getElementById('copybook-grid'),
            
            // 控制按钮
            printBtn: document.getElementById('print-btn'),
            exportPdfBtn: document.getElementById('export-pdf-btn'),
            exportWordBtn: document.getElementById('export-word-btn'),
            loadingOverlay: document.getElementById('loading-overlay'),
            
            // 设置面板元素
            gridSizeSlider: document.getElementById('grid-size-slider'),
            gridSizeValue: document.getElementById('grid-size-value'),
            charsPerRowSlider: document.getElementById('chars-per-row-slider'),
            charsPerRowValue: document.getElementById('chars-per-row-value'),
            rowSpacingSlider: document.getElementById('row-spacing-slider'),
            rowSpacingValue: document.getElementById('row-spacing-value'),
            fontFamilySelect: document.getElementById('font-family-select'),
            fontSizeSlider: document.getElementById('font-size-slider'),
            fontSizeValue: document.getElementById('font-size-value'),
            fontWeightSelect: document.getElementById('font-weight-select'),
            pinyinSizeSlider: document.getElementById('pinyin-size-slider'),
            pinyinSizeValue: document.getElementById('pinyin-size-value'),
            pageRowsSelect: document.getElementById('page-rows-select')
        };
        
        // 检查必要的元素是否存在
        const requiredElements = ['gradeSelect', 'generateBtn', 'copybookGrid'];
        for (const elementName of requiredElements) {
            if (!elements[elementName]) {
                console.error(`未找到必要元素: ${elementName}`);
                throw new Error(`无法找到元素: ${elementName}`);
            }
        }
        
        console.log('所有必要DOM元素已找到');
        
        // 当前字帖数据
        let currentCopybookData = [];
        let currentSource = ''; // 'system' 或 'excel'
        let currentGrade = '';
        
        // 字帖设置
        let settings = {
            gridType: 'tianzi', // 'tianzi' 或 'mizi'
            gridSize: 80, // 像素
            charsPerRow: 5,
            rowSpacing: 25, // 像素
            fontFamily: "'Ma Shan Zheng', cursive",
            fontSize: 48, // 像素
            fontWeight: 400,
            pinyinSize: 12, // 像素
            pageRows: 12,
            orientation: 'portrait', // 'portrait' 或 'landscape'
            showPinyin: true,
            showStrokes: true
        };
        
        // 初始化事件监听器
        function init() {
            console.log('初始化事件监听器...');
            
            // 年级选择事件
            elements.gradeSelect.addEventListener('change', function() {
                console.log('年级选择改变:', this.value);
                const selectedGrade = this.value;
                if (selectedGrade) {
                    loadGradeData(selectedGrade);
                    elements.generateBtn.disabled = false;
                    elements.generateBtn.innerHTML = '<i class="fas fa-magic"></i> 生成字帖';
                } else {
                    if (elements.characterCount) {
                        elements.characterCount.style.display = 'none';
                    }
                    elements.generateBtn.disabled = true;
                }
            });
            
            // 从系统数据生成字帖
            elements.generateBtn.addEventListener('click', function() {
                console.log('点击生成字帖按钮，当前数据量:', currentCopybookData.length);
                if (currentCopybookData.length > 0) {
                    generateCopybookFromSystem();
                } else {
                    alert('请先选择年级或加载数据');
                }
            });
            
            // 文件上传区域点击事件
            if (elements.uploadArea) {
                elements.uploadArea.addEventListener('click', function() {
                    console.log('点击上传区域');
                    if (elements.excelFileInput) {
                        elements.excelFileInput.click();
                    }
                });
            }
            
            // 文件选择事件
            if (elements.excelFileInput) {
                elements.excelFileInput.addEventListener('change', handleFileSelect);
            }
            
            // 从Excel生成字帖
            if (elements.generateExcelBtn) {
                elements.generateExcelBtn.addEventListener('click', function() {
                    console.log('点击从Excel生成字帖');
                    if (elements.excelFileInput && elements.excelFileInput.files.length > 0) {
                        processExcelFile(elements.excelFileInput.files[0]);
                    } else {
                        alert('请先选择Excel文件');
                    }
                });
            }
            
            // 打印按钮
            if (elements.printBtn) {
                elements.printBtn.addEventListener('click', function() {
                    window.print();
                });
            }
            
            // 导出PDF按钮
            if (elements.exportPdfBtn) {
                elements.exportPdfBtn.addEventListener('click', exportToPDF);
            }
            
            // 导出Word按钮
            if (elements.exportWordBtn) {
                elements.exportWordBtn.addEventListener('click', exportToWord);
            }
            
            // 设置面板事件监听
            setupSettingsPanel();
            
            // 加载年级数据
            loadAvailableGrades();
            
            console.log('事件监听器初始化完成');
        }
        
        // 设置面板事件监听
        function setupSettingsPanel() {
            console.log('设置面板事件监听...');
            
            // 网格类型按钮
            const gridTypeButtons = document.querySelectorAll('.grid-type-btn[data-type]');
            gridTypeButtons.forEach(button => {
                button.addEventListener('click', function() {
                    gridTypeButtons.forEach(btn => btn.classList.remove('active'));
                    this.classList.add('active');
                    settings.gridType = this.dataset.type;
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            });
            
            // 方格大小滑块
            if (elements.gridSizeSlider) {
                elements.gridSizeSlider.addEventListener('input', function() {
                    if (elements.gridSizeValue) {
                        elements.gridSizeValue.textContent = this.value;
                    }
                    settings.gridSize = parseInt(this.value);
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            }
            
            // 每行字数滑块
            if (elements.charsPerRowSlider) {
                elements.charsPerRowSlider.addEventListener('input', function() {
                    if (elements.charsPerRowValue) {
                        elements.charsPerRowValue.textContent = this.value;
                    }
                    settings.charsPerRow = parseInt(this.value);
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            }
            
            // 行间距滑块
            if (elements.rowSpacingSlider) {
                elements.rowSpacingSlider.addEventListener('input', function() {
                    if (elements.rowSpacingValue) {
                        elements.rowSpacingValue.textContent = this.value;
                    }
                    settings.rowSpacing = parseInt(this.value);
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            }
            
            // 字体类型选择
            if (elements.fontFamilySelect) {
                elements.fontFamilySelect.addEventListener('change', function() {
                    settings.fontFamily = this.value;
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            }
            
            // 字体大小滑块
            if (elements.fontSizeSlider) {
                elements.fontSizeSlider.addEventListener('input', function() {
                    if (elements.fontSizeValue) {
                        elements.fontSizeValue.textContent = this.value;
                    }
                    settings.fontSize = parseInt(this.value);
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            }
            
            // 字体粗细选择
            if (elements.fontWeightSelect) {
                elements.fontWeightSelect.addEventListener('change', function() {
                    settings.fontWeight = parseInt(this.value);
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            }
            
            // 拼音大小滑块
            if (elements.pinyinSizeSlider) {
                elements.pinyinSizeSlider.addEventListener('input', function() {
                    if (elements.pinyinSizeValue) {
                        elements.pinyinSizeValue.textContent = this.value;
                    }
                    settings.pinyinSize = parseInt(this.value);
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            }
            
            // 每页行数选择
            if (elements.pageRowsSelect) {
                elements.pageRowsSelect.addEventListener('change', function() {
                    settings.pageRows = parseInt(this.value);
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            }
            
            // 纸张方向按钮
            const orientationButtons = document.querySelectorAll('.grid-type-btn[data-orientation]');
            orientationButtons.forEach(button => {
                button.addEventListener('click', function() {
                    orientationButtons.forEach(btn => btn.classList.remove('active'));
                    this.classList.add('active');
                    settings.orientation = this.dataset.orientation;
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            });
            
            // 显示拼音按钮
            const showPinyinButtons = document.querySelectorAll('.grid-type-btn[data-show-pinyin]');
            showPinyinButtons.forEach(button => {
                button.addEventListener('click', function() {
                    showPinyinButtons.forEach(btn => btn.classList.remove('active'));
                    this.classList.add('active');
                    settings.showPinyin = this.dataset.showPinyin === 'true';
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            });
            
            // 显示笔画数按钮
            const showStrokesButtons = document.querySelectorAll('.grid-type-btn[data-show-strokes]');
            showStrokesButtons.forEach(button => {
                button.addEventListener('click', function() {
                    showStrokesButtons.forEach(btn => btn.classList.remove('active'));
                    this.classList.add('active');
                    settings.showStrokes = this.dataset.showStrokes === 'true';
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            });
        }
        
        // 加载可用的年级
        async function loadAvailableGrades() {
            console.log('开始加载可用年级...');
            try {
                const response = await fetch('data.json');
                if (!response.ok) {
                    throw new Error(`HTTP错误! 状态: ${response.status}`);
                }
                const data = await response.json();
                console.log('成功加载data.json');
                
                // 清空现有的选项（除了第一个）
                while (elements.gradeSelect.options.length > 1) {
                    elements.gradeSelect.remove(1);
                }
                
                // 填充年级选择
                const grades = data.map(item => item.grade);
                
                grades.forEach(grade => {
                    const option = document.createElement('option');
                    option.value = grade;
                    option.textContent = grade;
                    elements.gradeSelect.appendChild(option);
                });
                
                console.log('年级选择框已填充，选项数量:', elements.gradeSelect.options.length);
                
            } catch (error) {
                console.error('加载年级数据失败:', error);
                // 如果加载失败，添加一些默认选项
                const defaultGrades = [
                    '一年级上册', '一年级下册',
                    '二年级上册', '二年级下册',
                    '三年级上册', '三年级下册',
                    '四年级上册', '四年级下册',
                    '五年级上册', '五年级下册',
                    '六年级上册', '六年级下册'
                ];
                
                defaultGrades.forEach(grade => {
                    const option = document.createElement('option');
                    option.value = grade;
                    option.textContent = grade;
                    elements.gradeSelect.appendChild(option);
                });
                
                console.log('使用默认年级数据');
            }
        }
        
        // 加载年级数据
        async function loadGradeData(grade) {
            console.log('开始加载年级数据:', grade);
            try {
                const response = await fetch('data.json');
                const data = await response.json();
                
                const gradeData = data.find(item => item.grade === grade);
                
                if (gradeData && gradeData.characters) {
                    // 保存数据
                    currentCopybookData = gradeData.characters.map(item => ({
                        character: item.word,
                        pinyin: item.pinyin,
                        words: item.words || [],
                        strokeCount: estimateStrokeCount(item.word)
                    }));
                    
                    // 显示字数
                    if (elements.countNumber) {
                        elements.countNumber.textContent = currentCopybookData.length;
                    }
                    if (elements.characterCount) {
                        elements.characterCount.style.display = 'block';
                    }
                    
                    console.log(`已加载 ${grade} 的 ${currentCopybookData.length} 个生字`);
                } else {
                    console.warn(`未找到年级 ${grade} 的数据，使用示例数据`);
                    // 使用示例数据
                    currentCopybookData = [
                        {character: "天", pinyin: "tiān", words: ["天空", "今天"], strokeCount: 4},
                        {character: "地", pinyin: "dì", words: ["大地", "土地"], strokeCount: 6},
                        {character: "人", pinyin: "rén", words: ["人们", "好人"], strokeCount: 2},
                        {character: "你", pinyin: "nǐ", words: ["你好", "你们"], strokeCount: 7},
                        {character: "我", pinyin: "wǒ", words: ["我们", "自我"], strokeCount: 7},
                        {character: "他", pinyin: "tā", words: ["他们", "其他"], strokeCount: 5}
                    ];
                    if (elements.countNumber) {
                        elements.countNumber.textContent = currentCopybookData.length;
                    }
                    if (elements.characterCount) {
                        elements.characterCount.style.display = 'block';
                    }
                }
            } catch (error) {
                console.error('加载年级数据失败:', error);
                alert('加载数据失败，请检查网络连接');
            }
        }
        
        // 处理文件选择
        function handleFileSelect(event) {
            const file = event.target.files[0];
            if (file) {
                console.log('选择了文件:', file.name);
                // 显示文件信息
                if (elements.fileName) {
                    elements.fileName.textContent = file.name;
                }
                if (elements.fileInfo) {
                    elements.fileInfo.classList.add('show');
                }
                
                // 启用生成按钮
                if (elements.generateExcelBtn) {
                    elements.generateExcelBtn.disabled = false;
                    elements.generateExcelBtn.innerHTML = '<i class="fas fa-magic"></i> 从Excel生成字帖';
                }
            }
        }
        
        // 处理Excel文件
        function processExcelFile(file) {
            console.log('开始处理Excel文件:', file.name);
            if (elements.loadingOverlay) {
                elements.loadingOverlay.classList.add('show');
            }
            
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = e.target.result;
                    console.log('Excel文件读取成功');
                    
                    let workbook;
                    // 尝试不同的读取方式
                    try {
                        if (typeof data === 'string') {
                            workbook = XLSX.read(data, { type: 'binary' });
                        } else {
                            workbook = XLSX.read(data, { type: 'array' });
                        }
                    } catch (readError) {
                        console.error('读取Excel失败:', readError);
                        throw new Error('无法读取Excel文件，请确保文件格式正确');
                    }
                    
                    console.log('Excel工作表:', workbook.SheetNames);
                    
                    // 获取第一个工作表
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // 转换为JSON
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    console.log('Excel数据行数:', jsonData.length);
                    
                    if (jsonData.length === 0) {
                        throw new Error('Excel文件为空或格式不正确');
                    }
                    
                    currentCopybookData = [];
                    
                    // 处理每一行数据
                    jsonData.forEach((row, index) => {
                        // 查找"生字"列
                        let character = '';
                        
                        // 尝试不同的列名
                        if (row['生字'] !== undefined) {
                            character = String(row['生字']).trim();
                        } else if (row['word'] !== undefined) {
                            character = String(row['word']).trim();
                        } else if (row['汉字'] !== undefined) {
                            character = String(row['汉字']).trim();
                        } else if (row['字符'] !== undefined) {
                            character = String(row['字符']).trim();
                        } else if (row['字'] !== undefined) {
                            character = String(row['字']).trim();
                        } else {
                            // 如果没有找到标准列名，尝试第一列
                            const keys = Object.keys(row);
                            if (keys.length > 0) {
                                character = String(row[keys[0]]).trim();
                            }
                        }
                        
                        // 只处理非空的中文字符
                        if (character && /[\u4e00-\u9fa5]/.test(character)) {
                            // 获取拼音
                            let pinyin = '';
                            if (row['拼音'] !== undefined) {
                                pinyin = String(row['拼音']).trim();
                            } else if (row['pinyin'] !== undefined) {
                                pinyin = String(row['pinyin']).trim();
                            }
                            
                            // 获取组词
                            let words = [];
                            if (row['组词'] !== undefined) {
                                const wordsStr = String(row['组词']).trim();
                                if (wordsStr) {
                                    words = wordsStr.split(/[，,、\s]+/).filter(w => w.trim());
                                }
                            } else if (row['words'] !== undefined) {
                                const wordsStr = String(row['words']).trim();
                                if (wordsStr) {
                                    words = wordsStr.split(/[，,、\s]+/).filter(w => w.trim());
                                }
                            }
                            
                            currentCopybookData.push({
                                character: character,
                                pinyin: pinyin,
                                words: words,
                                strokeCount: estimateStrokeCount(character)
                            });
                        }
                    });
                    
                    console.log('处理完成，找到生字:', currentCopybookData.length);
                    
                    if (currentCopybookData.length > 0) {
                        currentSource = 'excel';
                        currentGrade = '自定义字帖';
                        generateCopybook('自定义字帖 - 生字练习');
                    } else {
                        if (elements.loadingOverlay) {
                            elements.loadingOverlay.classList.remove('show');
                        }
                        alert('Excel文件中未找到有效的生字数据。\n\n请确保：\n1. 文件包含"生字"列\n2. 生字列包含中文字符\n3. 文件格式正确（.xlsx或.xls）');
                    }
                    
                } catch (error) {
                    console.error('处理Excel文件失败:', error);
                    if (elements.loadingOverlay) {
                        elements.loadingOverlay.classList.remove('show');
                    }
                    alert(`处理Excel文件时出错：${error.message}\n\n请检查：\n1. 文件是否损坏\n2. 文件格式是否正确\n3. 是否包含"生字"列`);
                }
            };
            
            reader.onerror = function(error) {
                console.error('读取文件失败:', error);
                if (elements.loadingOverlay) {
                    elements.loadingOverlay.classList.remove('show');
                }
                alert('读取文件失败，请重试');
            };
            
            // 读取文件
            if (file.name.endsWith('.xlsx')) {
                reader.readAsArrayBuffer(file);
            } else {
                reader.readAsBinaryString(file);
            }
        }
        
        // 从系统数据生成字帖
        function generateCopybookFromSystem() {
            if (currentCopybookData.length === 0) {
                alert('请先选择年级');
                return;
            }
            
            currentSource = 'system';
            currentGrade = elements.gradeSelect.value;
            
            const title = `${currentGrade} - 生字练习字帖`;
            generateCopybook(title);
        }
        
        // 生成字帖
        function generateCopybook(title) {
            console.log('开始生成字帖，标题:', title, '数据量:', currentCopybookData.length);
            
            if (currentCopybookData.length === 0) {
                alert('没有可用的生字数据');
                return;
            }
            
            // 显示加载提示
            if (elements.loadingOverlay) {
                elements.loadingOverlay.classList.add('show');
            }
            
            // 使用setTimeout确保UI更新
            setTimeout(() => {
                try {
                    // 更新标题
                    if (elements.previewTitle) {
                        elements.previewTitle.textContent = title;
                    }
                    
                    // 生成字帖内容
                    generateCopybookContent();
                    
                    // 显示预览区域
                    if (elements.copybookPreview) {
                        elements.copybookPreview.classList.add('show');
                    }
                    
                    // 滚动到预览区域
                    if (elements.copybookPreview) {
                        elements.copybookPreview.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                    
                    console.log('字帖生成完成');
                    
                } catch (error) {
                    console.error('生成字帖失败:', error);
                    alert('生成字帖时出错，请重试');
                } finally {
                    // 确保隐藏加载提示
                    setTimeout(() => {
                        if (elements.loadingOverlay) {
                            elements.loadingOverlay.classList.remove('show');
                        }
                    }, 300);
                }
            }, 100);
        }
        
        // 生成字帖内容
        function generateCopybookContent() {
            console.log('生成字帖内容...');
            if (currentCopybookData.length === 0) return;
            
            try {
                // 清空内容
                elements.copybookGrid.innerHTML = '';
                
                // 创建字帖内容容器
                const copybookContainer = document.createElement('div');
                copybookContainer.className = 'copybook-content';
                copybookContainer.style.padding = '20px';
                copybookContainer.style.backgroundColor = '#fff';
                copybookContainer.style.borderRadius = '8px';
                
                // 计算每页字数
                const charsPerPage = settings.charsPerRow * settings.pageRows;
                const totalPages = Math.ceil(currentCopybookData.length / charsPerPage);
                
                // 创建分页
                for (let page = 0; page < totalPages; page++) {
                    const pageElement = document.createElement('div');
                    pageElement.className = 'copybook-page';
                    
                    // 添加页面标题（仅第一页）
                    if (page === 0) {
                        const titleElement = document.createElement('div');
                        titleElement.className = 'copybook-title';
                        titleElement.innerHTML = `
                            <h3>${elements.previewTitle ? elements.previewTitle.textContent : '字帖'}</h3>
                            <p>规范楷体硬笔字帖 - 共 ${currentCopybookData.length} 个生字 - 第 ${page + 1}/${totalPages} 页</p>
                            <p>日期: ${new Date().toLocaleDateString('zh-CN')}</p>
                        `;
                        pageElement.appendChild(titleElement);
                    }
                    
                    // 创建网格容器
                    const gridContainer = document.createElement('div');
                    gridContainer.className = 'copybook-grid-container';
                    gridContainer.style.display = 'grid';
                    gridContainer.style.gridTemplateColumns = `repeat(${settings.charsPerRow}, 1fr)`;
                    gridContainer.style.gap = '15px';
                    gridContainer.style.marginTop = '20px';
                    
                    // 计算当前页的生字
                    const startIndex = page * charsPerPage;
                    const endIndex = Math.min(startIndex + charsPerPage, currentCopybookData.length);
                    const pageCharacters = currentCopybookData.slice(startIndex, endIndex);
                    
                    // 添加每个生字的字帖单元格
                    pageCharacters.forEach((item, index) => {
                        // 每行的第一个单元格显示生字，其余为空白练习格
                        const isFirstInRow = (index % settings.charsPerRow) === 0;
                        
                        const characterCell = document.createElement('div');
                        characterCell.className = 'character-cell';
                        characterCell.style.width = `${settings.gridSize}px`;
                        characterCell.style.height = `${settings.gridSize}px`;
                        
                        // 添加网格背景
                        const gridBackground = document.createElement('div');
                        gridBackground.className = 'grid-background';
                        gridBackground.classList.add(`grid-${settings.gridType}`);
                        characterCell.appendChild(gridBackground);
                        
                        if (isFirstInRow) {
                            // 每行第一个格子显示生字
                            const characterContent = document.createElement('div');
                            characterContent.className = 'character-content';
                            characterContent.style.position = 'relative';
                            characterContent.style.zIndex = '1';
                            characterContent.style.width = '100%';
                            characterContent.style.height = '100%';
                            characterContent.style.display = 'flex';
                            characterContent.style.flexDirection = 'column';
                            characterContent.style.alignItems = 'center';
                            characterContent.style.justifyContent = 'center';
                            
                            // 添加生字
                            const characterElement = document.createElement('div');
                            characterElement.className = 'copybook-char';
                            characterElement.textContent = item.character;
                            characterElement.style.fontFamily = settings.fontFamily;
                            characterElement.style.fontSize = `${settings.fontSize}px`;
                            characterElement.style.fontWeight = settings.fontWeight;
                            characterElement.style.color = '#333';
                            characterElement.style.lineHeight = '1';
                            characterElement.style.textAlign = 'center';
                            characterContent.appendChild(characterElement);
                            
                            // 添加拼音（如果需要）
                            if (settings.showPinyin && item.pinyin) {
                                const pinyinElement = document.createElement('div');
                                pinyinElement.className = 'copybook-pinyin';
                                pinyinElement.textContent = item.pinyin;
                                pinyinElement.style.fontSize = `${settings.pinyinSize}px`;
                                pinyinElement.style.color = '#e74c3c';
                                pinyinElement.style.marginTop = '2px';
                                characterContent.appendChild(pinyinElement);
                            }
                            
                            // 添加笔画数（如果需要）
                            if (settings.showStrokes && item.strokeCount) {
                                const strokesElement = document.createElement('div');
                                strokesElement.className = 'copybook-strokes';
                                strokesElement.textContent = `笔画: ${item.strokeCount}`;
                                strokesElement.style.fontSize = '10px';
                                strokesElement.style.color = '#666';
                                strokesElement.style.marginTop = '2px';
                                characterContent.appendChild(strokesElement);
                            }
                            
                            characterCell.appendChild(characterContent);
                        } else {
                            // 其他格子为空白练习格
                            characterCell.innerHTML = '<div style="color:#999;font-size:12px;">练习格</div>';
                        }
                        
                        gridContainer.appendChild(characterCell);
                    });
                    
                    // 如果当前页不是满页，填充空白单元格
                    const remainingCells = charsPerPage - pageCharacters.length;
                    for (let i = 0; i < remainingCells; i++) {
                        const emptyCell = document.createElement('div');
                        emptyCell.className = 'character-cell';
                        emptyCell.style.width = `${settings.gridSize}px`;
                        emptyCell.style.height = `${settings.gridSize}px`;
                        
                        // 添加网格背景
                        const gridBackground = document.createElement('div');
                        gridBackground.className = 'grid-background';
                        gridBackground.classList.add(`grid-${settings.gridType}`);
                        emptyCell.appendChild(gridBackground);
                        
                        emptyCell.innerHTML = '<div style="color:#999;font-size:12px;">练习格</div>';
                        gridContainer.appendChild(emptyCell);
                    }
                    
                    pageElement.appendChild(gridContainer);
                    copybookContainer.appendChild(pageElement);
                    
                    // 如果不是最后一页，添加分页符
                    if (page < totalPages - 1) {
                        const pageBreak = document.createElement('div');
                        pageBreak.style.pageBreakAfter = 'always';
                        pageBreak.style.marginBottom = '50px';
                        copybookContainer.appendChild(pageBreak);
                    }
                }
                
                // 应用行间距
                const allRows = copybookContainer.querySelectorAll('.copybook-grid-container');
                allRows.forEach(row => {
                    row.style.marginBottom = `${settings.rowSpacing}px`;
                });
                
                // 添加内容到页面
                elements.copybookGrid.appendChild(copybookContainer);
                
                console.log('字帖内容生成完成');
                
            } catch (error) {
                console.error('生成字帖内容失败:', error);
                throw error;
            }
        }
        
        // 估算笔画数（简单版本）
        function estimateStrokeCount(character) {
            // 常见汉字笔画数映射
            const strokeMap = {
                '一': 1, '二': 2, '三': 3, '四': 5, '五': 4,
                '六': 4, '七': 2, '八': 2, '九': 2, '十': 2,
                '天': 4, '地': 6, '人': 2, '你': 7, '我': 7,
                '他': 5, '春': 9, '夏': 10, '秋': 9, '冬': 5,
                '日': 4, '月': 4, '水': 4, '火': 4, '土': 3,
                '木': 4, '金': 8, '山': 3, '石': 5, '田': 5,
                '上': 3, '下': 3, '中': 4, '大': 3, '小': 3,
                '学': 8, '生': 5, '字': 6, '文': 4, '语': 9
            };
            
            return strokeMap[character] || Math.floor(Math.random() * 10) + 3;
        }
        
        // 导出为PDF
        async function exportToPDF() {
            if (currentCopybookData.length === 0) {
                alert('请先生成字帖');
                return;
            }
            
            if (elements.loadingOverlay) {
                elements.loadingOverlay.classList.add('show');
            }
            
            try {
                // 使用html2canvas将字帖内容转换为图片
                const contentElement = elements.copybookGrid.querySelector('.copybook-content') || elements.copybookGrid;
                const canvas = await html2canvas(contentElement, {
                    scale: 2,
                    useCORS: true,
                    backgroundColor: '#ffffff',
                    logging: false
                });
                
                const imgData = canvas.toDataURL('image/png');
                
                // 创建PDF
                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF({
                    orientation: settings.orientation,
                    unit: 'mm',
                    format: 'a4'
                });
                
                const pageWidth = pdf.internal.pageSize.getWidth();
                const pageHeight = pdf.internal.pageSize.getHeight();
                
                // 计算图片尺寸
                const imgWidth = pageWidth - 20; // 左右各10mm边距
                const imgHeight = (canvas.height * imgWidth) / canvas.width;
                
                // 添加图片
                pdf.addImage(imgData, 'PNG', 10, 10, imgWidth, imgHeight);
                
                // 添加页眉
                pdf.setFontSize(12);
                pdf.setTextColor(74, 111, 165);
                const title = elements.previewTitle ? elements.previewTitle.textContent : '字帖';
                pdf.text(title, pageWidth / 2, 7, { align: 'center' });
                
                // 添加页脚
                pdf.setFontSize(10);
                pdf.setTextColor(100, 100, 100);
                pdf.text(`生成日期: ${new Date().toLocaleDateString('zh-CN')}`, 10, pageHeight - 10);
                pdf.text(`小学语文生字学习系统`, pageWidth / 2, pageHeight - 10, { align: 'center' });
                
                // 保存PDF
                const fileName = `${title.replace(/[<>:"/\\|?*]/g, '_')}.pdf`;
                pdf.save(fileName);
                
                console.log('PDF导出成功:', fileName);
                
            } catch (error) {
                console.error('导出PDF失败:', error);
                alert('导出PDF失败，请重试');
            } finally {
                if (elements.loadingOverlay) {
                    elements.loadingOverlay.classList.remove('show');
                }
            }
        }
        
        // 导出为Word
        async function exportToWord() {
            if (currentCopybookData.length === 0) {
                alert('请先生成字帖');
                return;
            }
            
            if (elements.loadingOverlay) {
                elements.loadingOverlay.classList.add('show');
            }
            
            try {
                // 创建Word文档内容
                const title = elements.previewTitle ? elements.previewTitle.textContent : '字帖';
                const htmlContent = `
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="UTF-8">
                        <title>${title}</title>
                        <style>
                            body { 
                                font-family: 'Microsoft YaHei', 'SimSun', sans-serif; 
                                margin: 20mm;
                                line-height: 1.5;
                            }
                            .copybook-title { 
                                text-align: center; 
                                margin-bottom: 30px; 
                            }
                            .copybook-title h3 { 
                                font-size: 24pt; 
                                color: #333; 
                                margin-bottom: 10px;
                            }
                            .copybook-title p { 
                                font-size: 12pt; 
                                color: #666;
                                margin: 5px 0;
                            }
                            .copybook-grid-container { 
                                display: grid; 
                                grid-template-columns: repeat(${settings.charsPerRow}, 1fr);
                                gap: 10px;
                                margin-bottom: ${settings.rowSpacing}px;
                            }
                            .character-cell { 
                                border: 1px solid #ccc;
                                width: ${settings.gridSize}px;
                                height: ${settings.gridSize}px;
                                position: relative;
                                display: flex;
                                align-items: center;
                                justify-content: center;
                            }
                            .character-content { 
                                text-align: center; 
                                z-index: 1;
                            }
                            .copybook-char { 
                                font-size: ${settings.fontSize}pt; 
                                margin-bottom: 5px; 
                                font-family: ${settings.fontFamily.replace(/'/g, '')};
                                font-weight: ${settings.fontWeight};
                            }
                            .copybook-pinyin { 
                                font-size: ${settings.pinyinSize}pt; 
                                color: #e74c3c; 
                            }
                            .copybook-strokes { 
                                font-size: 8pt; 
                                color: #666; 
                            }
                            @page { 
                                size: A4 ${settings.orientation}; 
                                margin: 20mm; 
                            }
                        </style>
                    </head>
                    <body>
                        <div class="copybook-title">
                            <h3>${title}</h3>
                            <p>规范楷体硬笔字帖 - 共 ${currentCopybookData.length} 个生字</p>
                            <p>日期: ${new Date().toLocaleDateString('zh-CN')}</p>
                        </div>
                        
                        <div style="text-align: center; margin-top: 30px; font-size: 10pt; color: #999;">
                            <p>生成日期: ${new Date().toLocaleDateString('zh-CN')}</p>
                            <p>小学语文生字学习系统 - 字帖打印功能</p>
                        </div>
                    </body>
                    </html>
                `;
                
                // 创建Blob并下载
                const blob = new Blob([htmlContent], { type: 'application/msword' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = `${title.replace(/[<>:"/\\|?*]/g, '_')}.doc`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                
                console.log('Word导出成功');
                
            } catch (error) {
                console.error('导出Word失败:', error);
                alert('导出Word失败，请重试');
            } finally {
                if (elements.loadingOverlay) {
                    elements.loadingOverlay.classList.remove('show');
                }
            }
        }
        
        // 初始化应用
        init();
        
    } catch (error) {
        console.error('初始化失败:', error);
        // 显示错误信息
        const errorMsg = document.createElement('div');
        errorMsg.style.cssText = `
            background-color: #f8d7da;
            color: #721c24;
            padding: 15px;
            margin: 20px;
            border-radius: 5px;
            border: 1px solid #f5c6cb;
        `;
        errorMsg.innerHTML = `
            <h3>页面初始化失败</h3>
            <p>错误: ${error.message}</p>
            <p>请检查HTML中是否存在以下元素：</p>
            <ul>
                <li>grade-select (年级选择)</li>
                <li>generate-btn (生成字帖按钮)</li>
                <li>copybook-grid (字帖网格容器)</li>
            </ul>
            <p>请刷新页面或检查控制台获取更多信息</p>
        `;
        document.body.insertBefore(errorMsg, document.body.firstChild);
    }
}

// 修复MetaMask冲突
function fixMetaMaskConflict() {
    // 检查是否是MetaMask扩展导致的错误
    if (window.ethereum && window.ethereum.isMetaMask) {
        console.log('检测到MetaMask扩展，正在尝试避免冲突...');
        // 尝试延迟执行以避免冲突
        setTimeout(main, 1000);
    } else {
        // 立即执行
        main();
    }
}

// 启动应用
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', fixMetaMaskConflict);
} else {
    fixMetaMaskConflict();
}