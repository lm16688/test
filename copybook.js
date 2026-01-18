// 字帖打印页面JavaScript
document.addEventListener('DOMContentLoaded', function() {
    console.log('页面加载完成，开始初始化...');
    
    // 获取DOM元素
    const gradeSelect = document.getElementById('grade-select');
    const generateBtn = document.getElementById('generate-btn');
    const characterCount = document.getElementById('character-count');
    const countNumber = document.getElementById('count-number');
    const uploadArea = document.getElementById('upload-area');
    const excelFileInput = document.getElementById('excel-file');
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const generateExcelBtn = document.getElementById('generate-excel-btn');
    const copybookPreview = document.getElementById('copybook-preview');
    const previewTitle = document.getElementById('preview-title');
    const copybookContent = document.getElementById('copybook-content');
    const printBtn = document.getElementById('print-btn');
    const exportPdfBtn = document.getElementById('export-pdf-btn');
    const exportWordBtn = document.getElementById('export-word-btn');
    const loadingOverlay = document.getElementById('loading-overlay');
    
    // 设置面板元素
    const gridTypeButtons = document.querySelectorAll('.grid-type-btn[data-type]');
    const gridSizeSlider = document.getElementById('grid-size-slider');
    const gridSizeValue = document.getElementById('grid-size-value');
    const charsPerRowSlider = document.getElementById('chars-per-row-slider');
    const charsPerRowValue = document.getElementById('chars-per-row-value');
    const rowSpacingSlider = document.getElementById('row-spacing-slider');
    const rowSpacingValue = document.getElementById('row-spacing-value');
    const fontFamilySelect = document.getElementById('font-family-select');
    const fontSizeSlider = document.getElementById('font-size-slider');
    const fontSizeValue = document.getElementById('font-size-value');
    const fontWeightSelect = document.getElementById('font-weight-select');
    const pinyinSizeSlider = document.getElementById('pinyin-size-slider');
    const pinyinSizeValue = document.getElementById('pinyin-size-value');
    const pageRowsSelect = document.getElementById('page-rows-select');
    const orientationButtons = document.querySelectorAll('.grid-type-btn[data-orientation]');
    const showPinyinButtons = document.querySelectorAll('.grid-type-btn[data-show-pinyin]');
    const showStrokesButtons = document.querySelectorAll('.grid-type-btn[data-show-strokes]');
    
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
        gradeSelect.addEventListener('change', function() {
            console.log('年级选择改变:', this.value);
            const selectedGrade = this.value;
            if (selectedGrade) {
                loadGradeData(selectedGrade);
                generateBtn.disabled = false;
                generateBtn.innerHTML = '<i class="fas fa-magic"></i> 生成字帖';
            } else {
                characterCount.style.display = 'none';
                generateBtn.disabled = true;
            }
        });
        
        // 从系统数据生成字帖
        generateBtn.addEventListener('click', function() {
            console.log('点击生成字帖按钮，当前数据量:', currentCopybookData.length);
            if (currentCopybookData.length > 0) {
                generateCopybookFromSystem();
            } else {
                alert('请先选择年级或加载数据');
            }
        });
        
        // 文件上传区域点击事件
        uploadArea.addEventListener('click', function() {
            console.log('点击上传区域');
            excelFileInput.click();
        });
        
        // 文件选择事件
        excelFileInput.addEventListener('change', handleFileSelect);
        
        // 从Excel生成字帖
        generateExcelBtn.addEventListener('click', function() {
            console.log('点击从Excel生成字帖，文件数量:', excelFileInput.files.length);
            if (excelFileInput.files.length > 0) {
                processExcelFile(excelFileInput.files[0]);
            } else {
                alert('请先选择Excel文件');
            }
        });
        
        // 打印按钮
        printBtn.addEventListener('click', function() {
            window.print();
        });
        
        // 导出PDF按钮
        exportPdfBtn.addEventListener('click', exportToPDF);
        
        // 导出Word按钮
        exportWordBtn.addEventListener('click', exportToWord);
        
        // 网格类型按钮
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
        gridSizeSlider.addEventListener('input', function() {
            gridSizeValue.textContent = this.value;
            settings.gridSize = parseInt(this.value);
            if (currentCopybookData.length > 0) {
                generateCopybookContent();
            }
        });
        
        // 每行字数滑块
        charsPerRowSlider.addEventListener('input', function() {
            charsPerRowValue.textContent = this.value;
            settings.charsPerRow = parseInt(this.value);
            if (currentCopybookData.length > 0) {
                generateCopybookContent();
            }
        });
        
        // 行间距滑块
        rowSpacingSlider.addEventListener('input', function() {
            rowSpacingValue.textContent = this.value;
            settings.rowSpacing = parseInt(this.value);
            if (currentCopybookData.length > 0) {
                generateCopybookContent();
            }
        });
        
        // 字体类型选择
        fontFamilySelect.addEventListener('change', function() {
            settings.fontFamily = this.value;
            if (currentCopybookData.length > 0) {
                generateCopybookContent();
            }
        });
        
        // 字体大小滑块
        fontSizeSlider.addEventListener('input', function() {
            fontSizeValue.textContent = this.value;
            settings.fontSize = parseInt(this.value);
            if (currentCopybookData.length > 0) {
                generateCopybookContent();
            }
        });
        
        // 字体粗细选择
        fontWeightSelect.addEventListener('change', function() {
            settings.fontWeight = parseInt(this.value);
            if (currentCopybookData.length > 0) {
                generateCopybookContent();
            }
        });
        
        // 拼音大小滑块
        pinyinSizeSlider.addEventListener('input', function() {
            pinyinSizeValue.textContent = this.value;
            settings.pinyinSize = parseInt(this.value);
            if (currentCopybookData.length > 0) {
                generateCopybookContent();
            }
        });
        
        // 每页行数选择
        pageRowsSelect.addEventListener('change', function() {
            settings.pageRows = parseInt(this.value);
            if (currentCopybookData.length > 0) {
                generateCopybookContent();
            }
        });
        
        // 纸张方向按钮
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
        
        // 加载年级数据
        loadAvailableGrades();
        
        console.log('事件监听器初始化完成');
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
            console.log('成功加载data.json，数据:', data);
            
            // 清空现有的选项（除了第一个）
            while (gradeSelect.options.length > 1) {
                gradeSelect.remove(1);
            }
            
            // 填充年级选择
            const grades = data.map(item => item.grade);
            console.log('提取的年级:', grades);
            
            grades.forEach(grade => {
                const option = document.createElement('option');
                option.value = grade;
                option.textContent = grade;
                gradeSelect.appendChild(option);
            });
            
            console.log('年级选择框已填充，选项数量:', gradeSelect.options.length);
            
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
                gradeSelect.appendChild(option);
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
                countNumber.textContent = currentCopybookData.length;
                characterCount.style.display = 'block';
                
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
                countNumber.textContent = currentCopybookData.length;
                characterCount.style.display = 'block';
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
            fileName.textContent = file.name;
            fileInfo.classList.add('show');
            
            // 启用生成按钮
            generateExcelBtn.disabled = false;
            generateExcelBtn.innerHTML = '<i class="fas fa-magic"></i> 从Excel生成字帖';
        }
    }
    
    // 处理Excel文件
    function processExcelFile(file) {
        console.log('开始处理Excel文件:', file.name);
        loadingOverlay.classList.add('show');
        
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
                        
                        console.log(`第${index + 1}行: 生字="${character}", 拼音="${pinyin}", 组词=${words.length}个`);
                    }
                });
                
                console.log('处理完成，找到生字:', currentCopybookData.length);
                
                if (currentCopybookData.length > 0) {
                    currentSource = 'excel';
                    currentGrade = '自定义字帖';
                    generateCopybook('自定义字帖 - 生字练习');
                } else {
                    loadingOverlay.classList.remove('show');
                    alert('Excel文件中未找到有效的生字数据。\n\n请确保：\n1. 文件包含"生字"列\n2. 生字列包含中文字符\n3. 文件格式正确（.xlsx或.xls）');
                }
                
            } catch (error) {
                console.error('处理Excel文件失败:', error);
                loadingOverlay.classList.remove('show');
                alert(`处理Excel文件时出错：${error.message}\n\n请检查：\n1. 文件是否损坏\n2. 文件格式是否正确\n3. 是否包含"生字"列`);
            }
        };
        
        reader.onerror = function(error) {
            console.error('读取文件失败:', error);
            loadingOverlay.classList.remove('show');
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
        currentGrade = gradeSelect.value;
        
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
        loadingOverlay.classList.add('show');
        
        // 使用setTimeout确保UI更新
        setTimeout(() => {
            try {
                // 更新标题
                previewTitle.textContent = title;
                
                // 生成字帖内容
                generateCopybookContent();
                
                // 显示预览区域
                copybookPreview.classList.add('show');
                
                // 滚动到预览区域
                copybookPreview.scrollIntoView({ behavior: 'smooth', block: 'start' });
                
                console.log('字帖生成完成');
                
            } catch (error) {
                console.error('生成字帖失败:', error);
                alert('生成字帖时出错，请重试');
            } finally {
                // 确保隐藏加载提示
                setTimeout(() => {
                    loadingOverlay.classList.remove('show');
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
            copybookContent.innerHTML = '';
            
            // 计算每页字数
            const charsPerPage = settings.charsPerRow * settings.pageRows;
            const totalPages = Math.ceil(currentCopybookData.length / charsPerPage);
            
            console.log('字帖设置:', settings);
            console.log('每页字数:', charsPerPage, '总页数:', totalPages);
            
            // 创建分页
            for (let page = 0; page < totalPages; page++) {
                const pageElement = document.createElement('div');
                pageElement.className = 'copybook-page';
                
                // 添加页面标题（仅第一页）
                if (page === 0) {
                    const titleElement = document.createElement('div');
                    titleElement.className = 'copybook-title';
                    titleElement.innerHTML = `
                        <h3>${previewTitle.textContent}</h3>
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
                
                console.log(`第${page + 1}页，生字 ${startIndex + 1}-${endIndex}，共${pageCharacters.length}个`);
                
                // 添加每个生字的字帖单元格
                pageCharacters.forEach((item, index) => {
                    // 每行的第一个单元格显示生字，其余为空白练习格
                    const isFirstInRow = (index % settings.charsPerRow) === 0;
                    
                    const characterCell = document.createElement('div');
                    characterCell.className = 'character-cell';
                    characterCell.style.width = `${settings.gridSize}px`;
                    characterCell.style.height = `${settings.gridSize}px`;
                    characterCell.style.minWidth = `${settings.gridSize}px`;
                    characterCell.style.minHeight = `${settings.gridSize}px`;
                    
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
                copybookContent.appendChild(pageElement);
                
                // 如果不是最后一页，添加分页符
                if (page < totalPages - 1) {
                    const pageBreak = document.createElement('div');
                    pageBreak.style.pageBreakAfter = 'always';
                    pageBreak.style.marginBottom = '50px';
                    copybookContent.appendChild(pageBreak);
                }
            }
            
            // 应用行间距
            const allRows = copybookContent.querySelectorAll('.copybook-grid-container');
            allRows.forEach(row => {
                row.style.marginBottom = `${settings.rowSpacing}px`;
            });
            
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
        
        loadingOverlay.classList.add('show');
        
        try {
            // 使用html2canvas将字帖内容转换为图片
            const content = document.getElementById('copybook-content');
            const canvas = await html2canvas(content, {
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
            let currentY = 10;
            pdf.addImage(imgData, 'PNG', 10, currentY, imgWidth, imgHeight);
            currentY += imgHeight + 10;
            
            // 添加页眉
            pdf.setFontSize(12);
            pdf.setTextColor(74, 111, 165);
            pdf.text(`${previewTitle.textContent}`, pageWidth / 2, 7, { align: 'center' });
            
            // 添加页脚
            pdf.setFontSize(10);
            pdf.setTextColor(100, 100, 100);
            pdf.text(`生成日期: ${new Date().toLocaleDateString('zh-CN')}`, 10, pageHeight - 10);
            pdf.text(`小学语文生字学习系统`, pageWidth / 2, pageHeight - 10, { align: 'center' });
            
            // 保存PDF
            const fileName = `${previewTitle.textContent.replace(/[<>:"/\\|?*]/g, '_')}.pdf`;
            pdf.save(fileName);
            
            console.log('PDF导出成功:', fileName);
            
        } catch (error) {
            console.error('导出PDF失败:', error);
            alert('导出PDF失败，请重试');
        } finally {
            loadingOverlay.classList.remove('show');
        }
    }
    
    // 导出为Word
    async function exportToWord() {
        if (currentCopybookData.length === 0) {
            alert('请先生成字帖');
            return;
        }
        
        loadingOverlay.classList.add('show');
        
        try {
            // 创建Word文档内容
            const htmlContent = `
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="UTF-8">
                    <title>${previewTitle.textContent}</title>
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
                        .grid-background {
                            position: absolute;
                            top: 0;
                            left: 0;
                            width: 100%;
                            height: 100%;
                            pointer-events: none;
                        }
                        ${settings.gridType === 'tianzi' ? `
                        .grid-tianzi {
                            background-image: 
                                linear-gradient(to right, #e0e0e0 1px, transparent 1px),
                                linear-gradient(to bottom, #e0e0e0 1px, transparent 1px),
                                linear-gradient(to right, #ddd 1px, transparent 1px),
                                linear-gradient(to bottom, #ddd 1px, transparent 1px);
                            background-size: 50% 100%, 100% 50%, 100% 100%, 100% 100%;
                            background-position: 50% 0, 0 50%, 0 0, 0 0;
                        }
                        ` : `
                        .grid-mizi {
                            background-image: 
                                linear-gradient(to right, #e0e0e0 1px, transparent 1px),
                                linear-gradient(to bottom, #e0e0e0 1px, transparent 1px),
                                linear-gradient(to right, #ddd 1px, transparent 1px),
                                linear-gradient(to bottom, #ddd 1px, transparent 1px),
                                linear-gradient(45deg, #eee 1px, transparent 1px),
                                linear-gradient(-45deg, #eee 1px, transparent 1px);
                            background-size: 50% 100%, 100% 50%, 100% 100%, 100% 100%, 100% 100%, 100% 100%;
                            background-position: 50% 0, 0 50%, 0 0, 0 0, 0 0, 0 0;
                        }
                        `}
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
                        <h3>${previewTitle.textContent}</h3>
                        <p>规范楷体硬笔字帖 - 共 ${currentCopybookData.length} 个生字</p>
                        <p>日期: ${new Date().toLocaleDateString('zh-CN')}</p>
                    </div>
                    
                    ${currentCopybookData.map((item, index) => {
                        const isFirstInRow = (index % settings.charsPerRow) === 0;
                        if (isFirstInRow) {
                            // 每行的开始，需要先关闭上一行（如果有）并开始新行
                            const rowStart = Math.floor(index / settings.charsPerRow) * settings.charsPerRow;
                            const rowData = currentCopybookData.slice(rowStart, rowStart + settings.charsPerRow);
                            
                            return `
                                <div class="copybook-grid-container">
                                    ${rowData.map((rowItem, rowIndex) => {
                                        if (rowIndex === 0) {
                                            // 第一个格子显示生字
                                            return `
                                                <div class="character-cell">
                                                    <div class="grid-background grid-${settings.gridType}"></div>
                                                    <div class="character-content">
                                                        <div class="copybook-char">${rowItem.character}</div>
                                                        ${settings.showPinyin && rowItem.pinyin ? `<div class="copybook-pinyin">${rowItem.pinyin}</div>` : ''}
                                                        ${settings.showStrokes && rowItem.strokeCount ? `<div class="copybook-strokes">笔画: ${rowItem.strokeCount}</div>` : ''}
                                                    </div>
                                                </div>
                                            `;
                                        } else {
                                            // 其他格子为空白
                                            return `
                                                <div class="character-cell">
                                                    <div class="grid-background grid-${settings.gridType}"></div>
                                                    <div style="color:#999;font-size:12px;">练习格</div>
                                                </div>
                                            `;
                                        }
                                    }).join('')}
                                </div>
                            `;
                        }
                        return '';
                    }).join('')}
                    
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
            link.download = `${previewTitle.textContent.replace(/[<>:"/\\|?*]/g, '_')}.doc`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            
            console.log('Word导出成功');
            
        } catch (error) {
            console.error('导出Word失败:', error);
            alert('导出Word失败，请重试');
        } finally {
            loadingOverlay.classList.remove('show');
        }
    }
    
    // 初始化应用
    init();
});