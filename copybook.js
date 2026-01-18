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
        
        // 获取DOM元素
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
        
        // 字帖设置 - 简化版
        let settings = {
            gridType: 'tianzi', // 'tianzi' 或 'mizi'
            gridSize: 80, // 像素
            charsPerRow: 5,
            rowSpacing: 25, // 像素
            fontFamily: "'Ma Shan Zheng', cursive",
            fontSize: 48, // 像素
            fontWeight: 400,
            pageRows: 12
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
           // 方格大小滑块
if (elements.gridSizeSlider) {
    elements.gridSizeSlider.addEventListener('input', function() {
        if (elements.gridSizeValue) {
            elements.gridSizeValue.textContent = this.value;
        }
        settings.gridSize = parseInt(this.value);
        if (currentCopybookData.length > 0) {
            // 重新生成内容时更新单元格高度
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
            
            // 每页行数选择
            if (elements.pageRowsSelect) {
                elements.pageRowsSelect.addEventListener('change', function() {
                    settings.pageRows = parseInt(this.value);
                    if (currentCopybookData.length > 0) {
                        generateCopybookContent();
                    }
                });
            }
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
        '学': 8, '生': 5, '字': 6, '文': 4, '语': 9,
        '红': 6, '花': 7, '草': 9, '树': 9, '林': 8,
        '河': 8, '海': 10, '云': 4, '雨': 8, '风': 4,
        '雷': 13, '电': 5, '光': 6, '明': 8, '亮': 9,
        '好': 6, '坏': 7, '美': 9, '丑': 4, '高': 10,
        '低': 7, '长': 4, '短': 12, '方': 4, '圆': 10,
        '东': 5, '南': 9, '西': 6, '北': 5, '中': 4,
        '左': 5, '右': 5, '前': 9, '后': 6, '上': 3,
        '下': 3, '里': 7, '外': 5, '多': 6, '少': 4,
        '大': 3, '小': 3, '远': 7, '近': 7, '快': 7,
        '慢': 14, '新': 13, '旧': 5, '老': 6, '幼': 5
    };
    
    // 如果字符在映射表中，返回对应的笔画数，否则返回估计值
    if (strokeMap[character]) {
        return strokeMap[character];
    }
    
    // 根据字符的Unicode范围估算笔画数
    const charCode = character.charCodeAt(0);
    
    // 简单估算规则：
    if (charCode >= 0x4E00 && charCode <= 0x4F00) {
        // 常用汉字区域
        return Math.floor(Math.random() * 5) + 3;
    } else if (charCode >= 0x3400 && charCode <= 0x4DBF) {
        // 扩展A区汉字
        return Math.floor(Math.random() * 8) + 5;
    } else {
        // 其他字符
        return Math.floor(Math.random() * 10) + 3;
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
        // 生成字帖内容 - 改为真正的田字格本子样式
function generateCopybookContent() {
    console.log('生成字帖内容...');
    if (currentCopybookData.length === 0) return;
    
    try {
        // 清空内容
        elements.copybookGrid.innerHTML = '';
        
        // 计算每页字数
        const charsPerRow = settings.charsPerRow;
        const pageRows = settings.pageRows;
        const charsPerPage = charsPerRow * pageRows;
        const totalPages = Math.ceil(currentCopybookData.length / charsPerPage);
        
        console.log('总页数:', totalPages, '每页字数:', charsPerPage);
        
        // 创建分页容器
        const paginationContainer = document.createElement('div');
        paginationContainer.className = 'pagination-container';
        paginationContainer.style.position = 'relative';
        paginationContainer.style.minHeight = '500px';
        
        // 创建分页内容
        const pagesContainer = document.createElement('div');
        pagesContainer.className = 'pages-container';
        pagesContainer.style.display = 'flex';
        pagesContainer.style.overflow = 'hidden';
        pagesContainer.style.position = 'relative';
        pagesContainer.style.width = '100%';
        
        // 创建分页
        for (let page = 0; page < totalPages; page++) {
            const pageElement = document.createElement('div');
            pageElement.className = 'copybook-page';
            pageElement.dataset.page = page + 1;
            pageElement.style.flex = '0 0 100%';
            pageElement.style.width = '100%';
            pageElement.style.padding = '30px';
            pageElement.style.backgroundColor = '#fff';
            pageElement.style.boxShadow = '0 2px 10px rgba(0,0,0,0.1)';
            pageElement.style.borderRadius = '8px';
            pageElement.style.marginBottom = '20px';
            pageElement.style.boxSizing = 'border-box';
            
            // 页面标题
            const titleElement = document.createElement('div');
            titleElement.className = 'page-title';
            titleElement.style.textAlign = 'center';
            titleElement.style.marginBottom = '25px';
            titleElement.style.paddingBottom = '15px';
            titleElement.style.borderBottom = '2px solid #f0f0f0';
            titleElement.innerHTML = `
                <h3 style="color:#4a6fa5;margin:0 0 8px 0;font-size:24px;">${elements.previewTitle ? elements.previewTitle.textContent : '字帖'}</h3>
                <p style="color:#666;margin:5px 0;font-size:14px;">规范楷体硬笔字帖 - 共 ${currentCopybookData.length} 个生字</p>
                <p style="color:#999;margin:5px 0;font-size:12px;">第 ${page + 1} 页 / 共 ${totalPages} 页 - 日期: ${new Date().toLocaleDateString('zh-CN')}</p>
            `;
            pageElement.appendChild(titleElement);
            
            // 创建田字格表格
            const tableContainer = document.createElement('div');
            tableContainer.className = 'table-container';
            tableContainer.style.width = '100%';
            
            // 计算当前页的生字
            const startIndex = page * charsPerPage;
            const endIndex = Math.min(startIndex + charsPerPage, currentCopybookData.length);
            const pageCharacters = currentCopybookData.slice(startIndex, endIndex);
            
            // 创建表格行
            for (let row = 0; row < pageRows; row++) {
                const tableRow = document.createElement('div');
                tableRow.className = 'table-row';
                tableRow.style.display = 'flex';
                tableRow.style.width = '100%';
                tableRow.style.borderTop = '1px solid #ccc';
                
                // 每行的最后一个单元格需要特殊处理
                if (row === pageRows - 1) {
                    tableRow.style.borderBottom = '1px solid #ccc';
                }
                
                // 创建表格单元格
                for (let col = 0; col < charsPerRow; col++) {
                    const tableCell = document.createElement('div');
                    tableCell.className = 'table-cell';
                    tableCell.style.flex = '1';
                    tableCell.style.position = 'relative';
                    tableCell.style.height = `${settings.gridSize}px`;
                    tableCell.style.borderRight = '1px solid #ccc';
                    tableCell.style.boxSizing = 'border-box';
                    
                    // 每行的最后一个单元格
                    if (col === charsPerRow - 1) {
                        tableCell.style.borderRight = 'none';
                    }
                    
                    // 创建田字格背景
                    const tianziGrid = document.createElement('div');
                    tianziGrid.className = 'tianzi-grid';
                    tianziGrid.style.position = 'absolute';
                    tianziGrid.style.top = '0';
                    tianziGrid.style.left = '0';
                    tianziGrid.style.width = '100%';
                    tianziGrid.style.height = '100%';
                    tianziGrid.style.pointerEvents = 'none';
                    
                    // 根据网格类型设置不同的背景
                    if (settings.gridType === 'tianzi') {
                        // 田字格：十字线
                        tianziGrid.style.cssText = `
                            position: absolute;
                            top: 0;
                            left: 0;
                            width: 100%;
                            height: 100%;
                            background-image: 
                                linear-gradient(to right, #ddd 1px, transparent 1px),
                                linear-gradient(to bottom, #ddd 1px, transparent 1px);
                            background-size: 50% 100%, 100% 50%;
                            background-position: 50% 0, 0 50%;
                            pointer-events: none;
                        `;
                    } else if (settings.gridType === 'mizi') {
                        // 米字格：十字线+对角线
                        tianziGrid.style.cssText = `
                            position: absolute;
                            top: 0;
                            left: 0;
                            width: 100%;
                            height: 100%;
                            background-image: 
                                linear-gradient(to right, #ddd 1px, transparent 1px),
                                linear-gradient(to bottom, #ddd 1px, transparent 1px),
                                linear-gradient(45deg, #eee 1px, transparent 1px),
                                linear-gradient(-45deg, #eee 1px, transparent 1px);
                            background-size: 50% 100%, 100% 50%, 100% 100%, 100% 100%;
                            background-position: 50% 0, 0 50%, 0 0, 0 0;
                            pointer-events: none;
                        `;
                    }
                    
                    tableCell.appendChild(tianziGrid);
                    
                    // 计算当前单元格对应的生字索引
                    const charIndex = row * charsPerRow + col;
                    const charData = pageCharacters[charIndex];
                    
                    if (charData && col === 0) {
                        // 每行第一个单元格显示生字
                        const characterDiv = document.createElement('div');
                        characterDiv.className = 'character-display';
                        characterDiv.style.position = 'absolute';
                        characterDiv.style.top = '50%';
                        characterDiv.style.left = '50%';
                        characterDiv.style.transform = 'translate(-50%, -50%)';
                        characterDiv.style.zIndex = '1';
                        characterDiv.style.fontFamily = settings.fontFamily;
                        characterDiv.style.fontSize = `${settings.fontSize}px`;
                        characterDiv.style.fontWeight = settings.fontWeight;
                        characterDiv.style.color = '#333';
                        characterDiv.style.textAlign = 'center';
                        characterDiv.textContent = charData.character;
                        
                        tableCell.appendChild(characterDiv);
                    }
                    
                    tableRow.appendChild(tableCell);
                }
                
                tableContainer.appendChild(tableRow);
            }
            
            pageElement.appendChild(tableContainer);
            pagesContainer.appendChild(pageElement);
        }
        
        // 添加分页控制
        if (totalPages > 1) {
            const paginationControls = document.createElement('div');
            paginationControls.className = 'pagination-controls';
            paginationControls.style.display = 'flex';
            paginationControls.style.justifyContent = 'center';
            paginationControls.style.alignItems = 'center';
            paginationControls.style.marginTop = '25px';
            paginationControls.style.gap = '15px';
            
            // 上一页按钮
            const prevButton = document.createElement('button');
            prevButton.className = 'page-btn prev-btn';
            prevButton.innerHTML = '<i class="fas fa-chevron-left"></i> 上一页';
            prevButton.disabled = true;
            prevButton.style.padding = '10px 20px';
            prevButton.style.backgroundColor = '#4a6fa5';
            prevButton.style.color = 'white';
            prevButton.style.border = 'none';
            prevButton.style.borderRadius = '6px';
            prevButton.style.cursor = 'pointer';
            prevButton.style.fontSize = '14px';
            
            // 页码显示
            const pageInfo = document.createElement('span');
            pageInfo.className = 'page-info';
            pageInfo.textContent = `第 1 页 / 共 ${totalPages} 页`;
            pageInfo.style.fontSize = '14px';
            pageInfo.style.color = '#666';
            
            // 下一页按钮
            const nextButton = document.createElement('button');
            nextButton.className = 'page-btn next-btn';
            nextButton.innerHTML = '下一页 <i class="fas fa-chevron-right"></i>';
            nextButton.disabled = totalPages <= 1;
            nextButton.style.padding = '10px 20px';
            nextButton.style.backgroundColor = totalPages > 1 ? '#4a6fa5' : '#cccccc';
            nextButton.style.color = 'white';
            nextButton.style.border = 'none';
            nextButton.style.borderRadius = '6px';
            nextButton.style.cursor = totalPages > 1 ? 'pointer' : 'not-allowed';
            nextButton.style.fontSize = '14px';
            
            // 分页按钮事件
            let currentPageIndex = 0;
            
            prevButton.addEventListener('click', function() {
                if (currentPageIndex > 0) {
                    currentPageIndex--;
                    updatePageDisplay();
                }
            });
            
            nextButton.addEventListener('click', function() {
                if (currentPageIndex < totalPages - 1) {
                    currentPageIndex++;
                    updatePageDisplay();
                }
            });
            
            function updatePageDisplay() {
                // 移动页面容器
                pagesContainer.style.transform = `translateX(-${currentPageIndex * 100}%)`;
                pagesContainer.style.transition = 'transform 0.3s ease';
                
                // 更新页码信息
                pageInfo.textContent = `第 ${currentPageIndex + 1} 页 / 共 ${totalPages} 页`;
                
                // 更新按钮状态
                prevButton.disabled = currentPageIndex === 0;
                nextButton.disabled = currentPageIndex === totalPages - 1;
                
                prevButton.style.backgroundColor = currentPageIndex === 0 ? '#cccccc' : '#4a6fa5';
                nextButton.style.backgroundColor = currentPageIndex === totalPages - 1 ? '#cccccc' : '#4a6fa5';
                prevButton.style.cursor = currentPageIndex === 0 ? 'not-allowed' : 'pointer';
                nextButton.style.cursor = currentPageIndex === totalPages - 1 ? 'not-allowed' : 'pointer';
            }
            
            paginationControls.appendChild(prevButton);
            paginationControls.appendChild(pageInfo);
            paginationControls.appendChild(nextButton);
            
            paginationContainer.appendChild(pagesContainer);
            paginationContainer.appendChild(paginationControls);
        } else {
            // 只有一页的情况
            paginationContainer.appendChild(pagesContainer);
        }
        
        // 添加内容到页面
        elements.copybookGrid.appendChild(paginationContainer);
        
        console.log('字帖内容生成完成');
        
    } catch (error) {
        console.error('生成字帖内容失败:', error);
        throw error;
    }
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
                // 创建PDF文档
                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF({
                    orientation: 'portrait',
                    unit: 'mm',
                    format: 'a4'
                });
                
                const pageWidth = pdf.internal.pageSize.getWidth();
                const pageHeight = pdf.internal.pageSize.getHeight();
                
                // 设置标题
                const title = elements.previewTitle ? elements.previewTitle.textContent : '字帖';
                pdf.setFontSize(16);
                pdf.setTextColor(74, 111, 165);
                pdf.text(title, pageWidth / 2, 20, { align: 'center' });
                
                // 设置副标题
                pdf.setFontSize(10);
                pdf.setTextColor(100, 100, 100);
                pdf.text(`规范楷体硬笔字帖 - 共 ${currentCopybookData.length} 个生字`, pageWidth / 2, 28, { align: 'center' });
                pdf.text(`日期: ${new Date().toLocaleDateString('zh-CN')}`, pageWidth / 2, 32, { align: 'center' });
                
                // 计算每页字数
                const charsPerRow = settings.charsPerRow;
                const pageRows = settings.pageRows;
                const charsPerPage = charsPerRow * pageRows;
                const totalPages = Math.ceil(currentCopybookData.length / charsPerPage);
                
                // 设置字体
                pdf.setFont('times', 'normal');
                
                let currentY = 45; // 开始Y坐标
                let currentPage = 0;
                
                // 处理所有生字
                for (let i = 0; i < currentCopybookData.length; i++) {
                    const item = currentCopybookData[i];
                    const rowInPage = Math.floor((i % charsPerPage) / charsPerRow);
                    const colInRow = (i % charsPerPage) % charsPerRow;
                    
                    // 如果是新的一页，添加新页面
                    if (i % charsPerPage === 0 && i > 0) {
                        pdf.addPage();
                        currentPage++;
                        currentY = 45;
                        
                        // 在新页面上添加标题
                        pdf.setFontSize(10);
                        pdf.text(`第 ${currentPage + 1}/${totalPages} 页`, pageWidth / 2, 15, { align: 'center' });
                    }
                    
                    // 计算每个格子的位置
                    const cellSize = 20; // 每个格子20mm
                    const startX = 10 + (colInRow * cellSize);
                    const cellY = currentY + (rowInPage * cellSize);
                    
                    // 绘制田字格或米字格
                    pdf.setDrawColor(200, 200, 200);
                    pdf.setLineWidth(0.1);
                    
                    // 绘制外边框
                    pdf.rect(startX, cellY, cellSize, cellSize);
                    
                    // 绘制田字格或米字格内部线条
                    if (settings.gridType === 'tianzi') {
                        // 田字格：十字线
                        pdf.line(startX, cellY + cellSize/2, startX + cellSize, cellY + cellSize/2);
                        pdf.line(startX + cellSize/2, cellY, startX + cellSize/2, cellY + cellSize);
                    } else if (settings.gridType === 'mizi') {
                        // 米字格：十字线+两条对角线
                        pdf.line(startX, cellY + cellSize/2, startX + cellSize, cellY + cellSize/2);
                        pdf.line(startX + cellSize/2, cellY, startX + cellSize/2, cellY + cellSize);
                        pdf.line(startX, cellY, startX + cellSize, cellY + cellSize);
                        pdf.line(startX + cellSize, cellY, startX, cellY + cellSize);
                    }
                    
                    // 如果是每行的第一个格子，写入生字
                    if (colInRow === 0) {
                        pdf.setFontSize(settings.fontSize * 0.35); // 调整字体大小
                        pdf.setTextColor(0, 0, 0);
                        pdf.text(item.character, startX + cellSize/2, cellY + cellSize/2 + 2, { align: 'center' });
                    }
                    
                    // 如果完成一行，更新Y坐标
                    if (colInRow === charsPerRow - 1) {
                        // 每行结束时增加行间距
                        if ((rowInPage + 1) % pageRows !== 0) {
                            // 不是页面的最后一行，继续下一行
                        }
                    }
                }
                
                // 添加页脚
                pdf.setFontSize(8);
                pdf.setTextColor(150, 150, 150);
                pdf.text(`小学语文生字学习系统`, pageWidth / 2, pageHeight - 10, { align: 'center' });
                pdf.text(`第 1 页`, 15, pageHeight - 10);
                pdf.text(`共 ${totalPages} 页`, pageWidth - 15, pageHeight - 10, { align: 'right' });
                
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
                // 计算每页字数
                const charsPerRow = settings.charsPerRow;
                const pageRows = settings.pageRows;
                const charsPerPage = charsPerRow * pageRows;
                const totalPages = Math.ceil(currentCopybookData.length / charsPerPage);
                
                // 创建Word文档内容
                const title = elements.previewTitle ? elements.previewTitle.textContent : '字帖';
                let wordContent = `
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
                            .page { 
                                margin-bottom: 30px;
                                page-break-after: always;
                            }
                            .title { 
                                text-align: center; 
                                margin-bottom: 20px; 
                            }
                            .title h3 { 
                                font-size: 20pt; 
                                color: #333; 
                                margin-bottom: 10px;
                            }
                            .title p { 
                                font-size: 11pt; 
                                color: #666;
                                margin: 3px 0;
                            }
                            .grid-container { 
                                display: grid; 
                                grid-template-columns: repeat(${charsPerRow}, 1fr);
                                gap: 0;
                                margin-bottom: ${settings.rowSpacing}px;
                            }
                            .character-cell { 
                                width: ${settings.gridSize * 0.35}mm;
                                height: ${settings.gridSize * 0.35}mm;
                                position: relative;
                                border: 1px solid #ccc;
                            }
                            .grid-lines {
                                position: absolute;
                                top: 0;
                                left: 0;
                                width: 100%;
                                height: 100%;
                                pointer-events: none;
                            }
                            .tianzi {
                                background-image: 
                                    linear-gradient(to right, #000 1px, transparent 1px),
                                    linear-gradient(to bottom, #000 1px, transparent 1px);
                                background-size: 50% 100%, 100% 50%;
                                background-position: 50% 0, 0 50%;
                            }
                            .mizi {
                                background-image: 
                                    linear-gradient(to right, #000 1px, transparent 1px),
                                    linear-gradient(to bottom, #000 1px, transparent 1px),
                                    linear-gradient(45deg, #888 1px, transparent 1px),
                                    linear-gradient(-45deg, #888 1px, transparent 1px);
                                background-size: 50% 100%, 100% 50%, 100% 100%, 100% 100%;
                                background-position: 50% 0, 0 50%, 0 0, 0 0;
                            }
                            .character-text {
                                position: absolute;
                                top: 50%;
                                left: 50%;
                                transform: translate(-50%, -50%);
                                font-size: ${settings.fontSize * 0.35}pt;
                                font-family: ${settings.fontFamily.replace(/'/g, '')};
                                font-weight: ${settings.fontWeight};
                                z-index: 1;
                            }
                            @page { 
                                size: A4 portrait; 
                                margin: 20mm; 
                            }
                        </style>
                    </head>
                    <body>
                `;
                
                // 分页处理
                for (let page = 0; page < totalPages; page++) {
                    wordContent += `<div class="page">`;
                    
                    // 页面标题
                    wordContent += `
                        <div class="title">
                            <h3>${title}</h3>
                            <p>规范楷体硬笔字帖 - 共 ${currentCopybookData.length} 个生字 - 第 ${page + 1}/${totalPages} 页</p>
                            <p>日期: ${new Date().toLocaleDateString('zh-CN')}</p>
                        </div>
                    `;
                    
                    // 计算当前页的生字
                    const startIndex = page * charsPerPage;
                    const endIndex = Math.min(startIndex + charsPerPage, currentCopybookData.length);
                    const pageCharacters = currentCopybookData.slice(startIndex, endIndex);
                    
                    // 创建网格
                    for (let row = 0; row < pageRows; row++) {
                        wordContent += `<div class="grid-container">`;
                        
                        for (let col = 0; col < charsPerRow; col++) {
                            const index = row * charsPerRow + col;
                            const item = pageCharacters[index];
                            
                            wordContent += `<div class="character-cell">`;
                            
                            // 网格背景
                            wordContent += `<div class="grid-lines ${settings.gridType}"></div>`;
                            
                            // 如果是每行的第一个格子，显示生字
                            if (col === 0 && item) {
                                wordContent += `<div class="character-text">${item.character}</div>`;
                            }
                            
                            wordContent += `</div>`;
                        }
                        
                        wordContent += `</div>`;
                    }
                    
                    wordContent += `</div>`;
                }
                
                // 页脚
                wordContent += `
                        <div style="text-align: center; margin-top: 30px; font-size: 9pt; color: #999;">
                            <p>生成日期: ${new Date().toLocaleDateString('zh-CN')}</p>
                            <p>小学语文生字学习系统 - 字帖打印功能</p>
                        </div>
                    </body>
                    </html>
                `;
                
                // 创建Blob并下载
                const blob = new Blob([wordContent], { type: 'application/msword' });
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