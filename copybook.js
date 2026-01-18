// 字帖打印页面JavaScript
document.addEventListener('DOMContentLoaded', function() {
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
    const copybookGrid = document.getElementById('copybook-grid');
    const refreshBtn = document.getElementById('refresh-btn');
    const printBtn = document.getElementById('print-btn');
    const applySettingsBtn = document.getElementById('apply-settings');
    
    // 设置控制元素
    const gridSizeSlider = document.getElementById('grid-size');
    const gridSizeValue = document.getElementById('grid-size-value');
    const gridGapSlider = document.getElementById('grid-gap');
    const gridGapValue = document.getElementById('grid-gap-value');
    const gridColorPicker = document.getElementById('grid-color');
    const gridTypeSelect = document.getElementById('grid-type');
    const fontFamilySelect = document.getElementById('font-family');
    const fontSizeSlider = document.getElementById('font-size');
    const fontSizeValue = document.getElementById('font-size-value');
    const fontWeightSlider = document.getElementById('font-weight');
    const fontWeightValue = document.getElementById('font-weight-value');
    const fontColorPicker = document.getElementById('font-color');
    const fontPreview = document.getElementById('font-preview');
    const showPinyinCheckbox = document.getElementById('show-pinyin');
    const showStrokeOrderCheckbox = document.getElementById('show-stroke-order');
    const showWordsCheckbox = document.getElementById('show-words');
    const showStrokeCountCheckbox = document.getElementById('show-stroke-count');
    const charsPerRowSelect = document.getElementById('chars-per-row');
    
    // 当前字帖数据
    let currentCopybookData = [];
    let currentSource = ''; // 'system' 或 'excel'
    let currentGrade = '';
    
    // 当前设置
    let currentSettings = {
        gridSize: 150,
        gridGap: 20,
        gridColor: '#cccccc',
        gridType: 'square',
        fontFamily: "'Ma Shan Zheng', cursive",
        fontSize: 48,
        fontWeight: 400,
        fontColor: '#000000',
        showPinyin: true,
        showStrokeOrder: true,
        showWords: true,
        showStrokeCount: false,
        charsPerRow: 3
    };
    
    // 初始化事件监听器
    function init() {
        // 初始化设置值显示
        updateSettingsDisplay();
        
        // 年级选择事件
        gradeSelect.addEventListener('change', function() {
            const selectedGrade = this.value;
            if (selectedGrade) {
                loadGradeData(selectedGrade);
                generateBtn.disabled = false;
            } else {
                characterCount.style.display = 'none';
                generateBtn.disabled = true;
            }
        });
        
        // 从系统数据生成字帖
        generateBtn.addEventListener('click', function() {
            if (currentCopybookData.length > 0) {
                generateCopybookFromSystem();
            }
        });
        
        // 文件上传区域点击事件
        uploadArea.addEventListener('click', function() {
            excelFileInput.click();
        });
        
        // 文件选择事件
        excelFileInput.addEventListener('change', handleFileSelect);
        
        // 从Excel生成字帖
        generateExcelBtn.addEventListener('click', function() {
            if (excelFileInput.files.length > 0) {
                processExcelFile(excelFileInput.files[0]);
            }
        });
        
        // 重新生成按钮
        refreshBtn.addEventListener('click', function() {
            if (currentSource === 'system') {
                generateCopybookFromSystem();
            } else if (currentSource === 'excel') {
                if (excelFileInput.files.length > 0) {
                    processExcelFile(excelFileInput.files[0]);
                }
            }
        });
        
        // 打印按钮
        printBtn.addEventListener('click', function() {
            window.print();
        });
        
        // 应用设置按钮
        applySettingsBtn.addEventListener('click', applySettings);
        
        // 设置滑块和选择器事件
        setupSettingsEvents();
        
        // 加载年级数据
        loadAvailableGrades();
    }
    
    // 设置滑块和选择器事件
    function setupSettingsEvents() {
        // 方格大小
        gridSizeSlider.addEventListener('input', function() {
            gridSizeValue.textContent = this.value + 'px';
            currentSettings.gridSize = parseInt(this.value);
        });
        
        // 行间距
        gridGapSlider.addEventListener('input', function() {
            gridGapValue.textContent = this.value + 'px';
            currentSettings.gridGap = parseInt(this.value);
        });
        
        // 网格线颜色
        gridColorPicker.addEventListener('input', function() {
            currentSettings.gridColor = this.value;
        });
        
        // 网格类型
        gridTypeSelect.addEventListener('change', function() {
            currentSettings.gridType = this.value;
        });
        
        // 字体类型
        fontFamilySelect.addEventListener('change', function() {
            currentSettings.fontFamily = this.value;
            updateFontPreview();
        });
        
        // 字体大小
        fontSizeSlider.addEventListener('input', function() {
            fontSizeValue.textContent = this.value + 'px';
            currentSettings.fontSize = parseInt(this.value);
            updateFontPreview();
        });
        
        // 字体粗细
        fontWeightSlider.addEventListener('input', function() {
            const weight = parseInt(this.value);
            let weightText = '';
            if (weight <= 300) weightText = '细';
            else if (weight <= 400) weightText = '正常';
            else if (weight <= 600) weightText = '中等';
            else if (weight <= 700) weightText = '加粗';
            else weightText = '特粗';
            
            fontWeightValue.textContent = weightText;
            currentSettings.fontWeight = weight;
            updateFontPreview();
        });
        
        // 字体颜色
        fontColorPicker.addEventListener('input', function() {
            currentSettings.fontColor = this.value;
            updateFontPreview();
        });
        
        // 显示选项
        showPinyinCheckbox.addEventListener('change', function() {
            currentSettings.showPinyin = this.checked;
        });
        
        showStrokeOrderCheckbox.addEventListener('change', function() {
            currentSettings.showStrokeOrder = this.checked;
        });
        
        showWordsCheckbox.addEventListener('change', function() {
            currentSettings.showWords = this.checked;
        });
        
        showStrokeCountCheckbox.addEventListener('change', function() {
            currentSettings.showStrokeCount = this.checked;
        });
        
        // 每行字数
        charsPerRowSelect.addEventListener('change', function() {
            currentSettings.charsPerRow = parseInt(this.value);
        });
    }
    
    // 更新设置显示
    function updateSettingsDisplay() {
        gridSizeValue.textContent = currentSettings.gridSize + 'px';
        gridSizeSlider.value = currentSettings.gridSize;
        gridGapValue.textContent = currentSettings.gridGap + 'px';
        gridGapSlider.value = currentSettings.gridGap;
        gridColorPicker.value = currentSettings.gridColor;
        gridTypeSelect.value = currentSettings.gridType;
        fontFamilySelect.value = currentSettings.fontFamily;
        fontSizeValue.textContent = currentSettings.fontSize + 'px';
        fontSizeSlider.value = currentSettings.fontSize;
        
        // 字体粗细显示
        let weightText = '';
        if (currentSettings.fontWeight <= 300) weightText = '细';
        else if (currentSettings.fontWeight <= 400) weightText = '正常';
        else if (currentSettings.fontWeight <= 600) weightText = '中等';
        else if (currentSettings.fontWeight <= 700) weightText = '加粗';
        else weightText = '特粗';
        
        fontWeightValue.textContent = weightText;
        fontWeightSlider.value = currentSettings.fontWeight;
        fontColorPicker.value = currentSettings.fontColor;
        showPinyinCheckbox.checked = currentSettings.showPinyin;
        showStrokeOrderCheckbox.checked = currentSettings.showStrokeOrder;
        showWordsCheckbox.checked = currentSettings.showWords;
        showStrokeCountCheckbox.checked = currentSettings.showStrokeCount;
        charsPerRowSelect.value = currentSettings.charsPerRow;
        
        updateFontPreview();
    }
    
    // 更新字体预览
    function updateFontPreview() {
        fontPreview.style.fontFamily = currentSettings.fontFamily;
        fontPreview.style.fontSize = currentSettings.fontSize + 'px';
        fontPreview.style.fontWeight = currentSettings.fontWeight;
        fontPreview.style.color = currentSettings.fontColor;
    }
    
    // 应用设置
    function applySettings() {
        if (currentCopybookData.length > 0) {
            // 重新生成字帖
            if (currentSource === 'system') {
                generateCopybookFromSystem();
            } else if (currentSource === 'excel') {
                if (excelFileInput.files.length > 0) {
                    processExcelFile(excelFileInput.files[0]);
                }
            }
        }
    }
    
    // 加载可用的年级
    async function loadAvailableGrades() {
        try {
            const response = await fetch('data.json');
            const data = await response.json();
            
            // 填充年级选择
            const grades = data.map(item => item.grade);
            
            // 如果某些年级不在选择列表中，可以动态添加
            grades.forEach(grade => {
                const option = document.createElement('option');
                option.value = grade;
                option.textContent = grade;
                gradeSelect.appendChild(option);
            });
            
        } catch (error) {
            console.error('加载年级数据失败:', error);
        }
    }
    
    // 加载年级数据
    async function loadGradeData(grade) {
        try {
            const response = await fetch('data.json');
            const data = await response.json();
            
            const gradeData = data.find(item => item.grade === grade);
            
            if (gradeData && gradeData.characters) {
                // 保存数据
                currentCopybookData = gradeData.characters.map(item => ({
                    character: item.word,
                    pinyin: item.pinyin,
                    words: item.words || []
                }));
                
                // 显示字数
                countNumber.textContent = currentCopybookData.length;
                characterCount.style.display = 'block';
                
                console.log(`已加载 ${grade} 的 ${currentCopybookData.length} 个生字`);
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
            // 显示文件信息
            fileName.textContent = file.name;
            fileInfo.classList.add('show');
            
            // 启用生成按钮
            generateExcelBtn.disabled = false;
        }
    }
    
    // 处理Excel文件
    function processExcelFile(file) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'binary' });
                
                // 获取第一个工作表
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // 将工作表转换为JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                // 处理Excel数据
                processExcelData(jsonData);
                
            } catch (error) {
                console.error('处理Excel文件时出错:', error);
                alert('处理Excel文件时出错，请确保文件格式正确');
            }
        };
        
        reader.onerror = function() {
            alert('读取文件失败');
        };
        
        reader.readAsBinaryString(file);
    }
    
    // 处理Excel数据
    function processExcelData(data) {
        currentCopybookData = [];
        
        // 假设Excel第一行是标题行，数据从第二行开始
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row && row.length > 0) {
                const character = row[0]?.toString().trim();
                
                if (character && character.length === 1) {
                    // 获取拼音（如果有）
                    const pinyin = row.length > 1 ? row[1]?.toString().trim() : '';
                    
                    // 获取组词（如果有）
                    const words = row.length > 2 ? row[2]?.toString().trim().split(/[,，]/) : [];
                    
                    currentCopybookData.push({
                        character: character,
                        pinyin: pinyin,
                        words: words.filter(word => word.trim() !== '')
                    });
                }
            }
        }
        
        if (currentCopybookData.length > 0) {
            currentSource = 'excel';
            generateCopybookFromExcel();
        } else {
            alert('Excel文件中没有找到有效的汉字数据');
        }
    }
    
    // 从系统数据生成字帖
    function generateCopybookFromSystem() {
        if (currentCopybookData.length === 0) {
            alert('请先选择年级或上传Excel文件');
            return;
        }
        
        currentSource = 'system';
        renderCopybook();
    }
    
    // 从Excel数据生成字帖
    function generateCopybookFromExcel() {
        if (currentCopybookData.length === 0) {
            alert('Excel文件中没有有效的汉字数据');
            return;
        }
        
        renderCopybook();
    }
    
    // 渲染字帖
    function renderCopybook() {
        // 清空预览区
        copybookGrid.innerHTML = '';
        
        // 设置标题
        if (currentSource === 'system') {
            previewTitle.textContent = `${gradeSelect.options[gradeSelect.selectedIndex].text} 生字字帖`;
        } else {
            previewTitle.textContent = `自定义字帖 (${currentCopybookData.length}个汉字)`;
        }
        
        // 显示预览区
        copybookPreview.classList.add('show');
        
        // 创建网格容器
        const gridContainer = document.createElement('div');
        gridContainer.className = 'grid-container';
        
        // 设置网格样式
        gridContainer.style.gridTemplateColumns = `repeat(${currentSettings.charsPerRow}, 1fr)`;
        gridContainer.style.gap = `${currentSettings.gridGap}px`;
        
        // 添加每个汉字
        currentCopybookData.forEach((item, index) => {
            const gridItem = createGridItem(item, index);
            gridContainer.appendChild(gridItem);
        });
        
        copybookGrid.appendChild(gridContainer);
    }
    
    // 创建单个网格项
    function createGridItem(item, index) {
        const gridItem = document.createElement('div');
        gridItem.className = 'grid-item';
        
        // 设置网格大小
        gridItem.style.width = `${currentSettings.gridSize}px`;
        gridItem.style.height = `${currentSettings.gridSize}px`;
        
        // 创建网格内部结构
        const gridInner = document.createElement('div');
        gridInner.className = `grid-inner ${currentSettings.gridType}`;
        
        // 设置网格线颜色
        if (currentSettings.gridType === 'square') {
            gridInner.style.borderColor = currentSettings.gridColor;
        } else if (currentSettings.gridType === 'tian') {
            gridInner.style.borderColor = currentSettings.gridColor;
            gridInner.style.backgroundColor = 'transparent';
        }
        
        // 添加汉字
        const characterElement = document.createElement('div');
        characterElement.className = 'character';
        characterElement.textContent = item.character;
        characterElement.style.fontFamily = currentSettings.fontFamily;
        characterElement.style.fontSize = `${currentSettings.fontSize}px`;
        characterElement.style.fontWeight = currentSettings.fontWeight;
        characterElement.style.color = currentSettings.fontColor;
        
        gridInner.appendChild(characterElement);
        
        // 添加拼音（如果启用）
        if (currentSettings.showPinyin && item.pinyin) {
            const pinyinElement = document.createElement('div');
            pinyinElement.className = 'pinyin';
            pinyinElement.textContent = item.pinyin;
            gridInner.appendChild(pinyinElement);
        }
        
        // 添加笔画顺序（如果启用）
        if (currentSettings.showStrokeOrder) {
            const strokeOrderElement = document.createElement('div');
            strokeOrderElement.className = 'stroke-order';
            strokeOrderElement.textContent = `${index + 1}`;
            gridInner.appendChild(strokeOrderElement);
        }
        
        // 添加笔画数（如果启用）
        if (currentSettings.showStrokeCount) {
            // 这里使用简单的笔画数估算，实际项目中可能需要完整的笔画数据库
            const strokeCount = estimateStrokeCount(item.character);
            const strokeCountElement = document.createElement('div');
            strokeCountElement.className = 'stroke-count';
            strokeCountElement.textContent = `${strokeCount}画`;
            gridInner.appendChild(strokeCountElement);
        }
        
        // 添加组词（如果启用且有组词）
        if (currentSettings.showWords && item.words && item.words.length > 0) {
            const wordsElement = document.createElement('div');
            wordsElement.className = 'words';
            wordsElement.textContent = item.words.slice(0, 2).join(' ');
            gridInner.appendChild(wordsElement);
        }
        
        gridItem.appendChild(gridInner);
        return gridItem;
    }
    
    // 估算汉字笔画数（简化版）
    function estimateStrokeCount(character) {
        // 常见汉字笔画数映射（简化版，实际项目需要完整数据库）
        const strokeMap = {
            '一': 1, '二': 2, '三': 3, '人': 2, '大': 3,
            '天': 4, '地': 6, '日': 4, '月': 4, '水': 4,
            '火': 4, '木': 4, '金': 8, '土': 3, '上': 3,
            '下': 3, '中': 4, '国': 8, '学': 8, '生': 5,
            '我': 7, '你': 7, '他': 5, '好': 6, '爱': 10,
            '家': 10, '学': 8, '校': 10, '老': 6, '师': 6
        };
        
        return strokeMap[character] || 8; // 默认返回8画
    }
    
    // 添加打印样式
    function addPrintStyles() {
        const printStyles = document.createElement('style');
        printStyles.textContent = `
            @media print {
                .header, .settings-panel, .file-upload, .controls {
                    display: none !important;
                }
                
                .preview-section {
                    width: 100% !important;
                    margin: 0 !important;
                    padding: 0 !important;
                    border: none !important;
                    box-shadow: none !important;
                }
                
                .copybook-preview {
                    display: block !important;
                    width: 100% !important;
                    max-width: none !important;
                }
                
                .preview-header {
                    margin-bottom: 20px !important;
                }
                
                .grid-container {
                    break-inside: avoid;
                    page-break-inside: avoid;
                    grid-template-columns: repeat(${currentSettings.charsPerRow}, 1fr) !important;
                }
                
                .grid-item {
                    width: ${currentSettings.gridSize}px !important;
                    height: ${currentSettings.gridSize}px !important;
                }
                
                body {
                    margin: 0 !important;
                    padding: 20px !important;
                    font-family: Arial, sans-serif !important;
                }
                
                * {
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                }
                
                @page {
                    margin: 0.5cm;
                    size: A4;
                }
            }
            
            /* 田字格特殊样式 */
            .grid-inner.tian {
                position: relative;
                border: 1px solid;
            }
            
            .grid-inner.tian::before,
            .grid-inner.tian::after {
                content: '';
                position: absolute;
                background-color: var(--grid-color, #cccccc);
            }
            
            .grid-inner.tian::before {
                top: 0;
                left: 50%;
                width: 1px;
                height: 100%;
                transform: translateX(-50%);
            }
            
            .grid-inner.tian::after {
                top: 50%;
                left: 0;
                width: 100%;
                height: 1px;
                transform: translateY(-50%);
            }
            
            /* 网格项内部元素样式 */
            .grid-inner {
                position: relative;
                display: flex;
                justify-content: center;
                align-items: center;
                width: 100%;
                height: 100%;
                border: 1px solid;
            }
            
            .character {
                font-size: var(--font-size, 48px);
                font-weight: var(--font-weight, 400);
                color: var(--font-color, #000000);
                font-family: var(--font-family, "'Ma Shan Zheng', cursive");
                z-index: 1;
            }
            
            .pinyin {
                position: absolute;
                top: 5px;
                left: 0;
                right: 0;
                text-align: center;
                font-size: 14px;
                color: #666;
            }
            
            .stroke-order {
                position: absolute;
                top: 5px;
                left: 5px;
                font-size: 12px;
                color: #999;
            }
            
            .stroke-count {
                position: absolute;
                bottom: 5px;
                right: 5px;
                font-size: 12px;
                color: #999;
            }
            
            .words {
                position: absolute;
                bottom: 5px;
                left: 0;
                right: 0;
                text-align: center;
                font-size: 12px;
                color: #666;
            }
        `;
        document.head.appendChild(printStyles);
    }
    
    // 导出字帖为PDF
    function exportToPDF() {
        // 这个功能需要html2canvas和jsPDF库
        // 这里只提供基本思路
        console.log('PDF导出功能需要额外库支持');
        
        // 示例代码：
        /*
        html2canvas(copybookPreview).then(canvas => {
            const imgData = canvas.toDataURL('image/png');
            const pdf = new jsPDF('p', 'mm', 'a4');
            const imgWidth = 210;
            const pageHeight = 295;
            const imgHeight = canvas.height * imgWidth / canvas.width;
            let heightLeft = imgHeight;
            let position = 0;
            
            pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;
            
            while (heightLeft >= 0) {
                position = heightLeft - imgHeight;
                pdf.addPage();
                pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
                heightLeft -= pageHeight;
            }
            
            pdf.save('字帖.pdf');
        });
        */
    }
    
    // 添加导出功能
    function addExportButton() {
        const exportBtn = document.createElement('button');
        exportBtn.id = 'export-btn';
        exportBtn.className = 'btn btn-secondary';
        exportBtn.innerHTML = '<i class="fas fa-file-pdf"></i> 导出PDF';
        exportBtn.addEventListener('click', exportToPDF);
        
        const controls = document.querySelector('.controls');
        if (controls) {
            controls.appendChild(exportBtn);
        }
    }
    
    // 添加CSS变量更新
    function updateCSSVariables() {
        const root = document.documentElement;
        root.style.setProperty('--grid-color', currentSettings.gridColor);
        root.style.setProperty('--font-family', currentSettings.fontFamily);
        root.style.setProperty('--font-size', currentSettings.fontSize + 'px');
        root.style.setProperty('--font-weight', currentSettings.fontWeight);
        root.style.setProperty('--font-color', currentSettings.fontColor);
    }
    
    // 初始化页面
    init();
    addPrintStyles();
    updateCSSVariables();
    // 如果需要导出功能，取消下面注释
    // addExportButton();
    
    // 监听设置变化更新CSS变量
    applySettingsBtn.addEventListener('click', updateCSSVariables);
});