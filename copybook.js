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
            updateCharsPerRow();
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
        
        // 根据方格大小自动调整每行字数
        function updateCharsPerRow() {
            const gridSize = parseInt(gridSizeSlider.value);
            let charsPerRow;
            
            if (gridSize <= 70) {
                charsPerRow = 8;
            } else if (gridSize <= 80) {
                charsPerRow = 6;
            } else if (gridSize <= 90) {
                charsPerRow = 5;
            } else if (gridSize <= 100) {
                charsPerRow = 4;
            } else {
                charsPerRow = 3;
            }
            
            charsPerRowSlider.value = charsPerRow;
            charsPerRowValue.textContent = charsPerRow;
            settings.charsPerRow = charsPerRow;
        }
        
        // 加载年级数据
        loadAvailableGrades();
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
                // 检查是否已存在
                const existingOption = Array.from(gradeSelect.options).find(opt => opt.value === grade);
                if (!existingOption) {
                    const option = document.createElement('option');
                    option.value = grade;
                    option.textContent = grade;
                    gradeSelect.appendChild(option);
                }
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
                    words: item.words || [],
                    strokeCount: estimateStrokeCount(item.word)
                }));
                
                // 显示字数
                countNumber.textContent = currentCopybookData.length;
                characterCount.style.display = 'block';
                
                console.log(`已加载 ${grade} 的 ${currentCopybookData.length} 个生字`);
            }
        } catch (error) {
            console.error('加载年级数据失败:', error);
            alert('加载数据失败，请检查网络连接或data.json文件');
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
        loadingOverlay.classList.add('show');
        
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                console.log('开始处理Excel文件，大小:', data.length);
                
                // 使用正确的读取方式
                let workbook;
                if (typeof data === 'string') {
                    // 二进制字符串
                    workbook = XLSX.read(data, { type: 'binary' });
                } else {
                    // ArrayBuffer
                    workbook = XLSX.read(data, { type: 'array' });
                }
                
                console.log('Excel工作表:', workbook.SheetNames);
                
                // 获取第一个工作表
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // 转换为JSON - 使用原始模式获取所有数据
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                console.log('Excel数据行数:', jsonData.length);
                
                // 处理数据 - 找到"生字"列的索引
                let characterColIndex = -1;
                let pinyinColIndex = -1;
                let wordsColIndex = -1;
                
                // 查找表头行
                if (jsonData.length > 0) {
                    const headerRow = jsonData[0];
                    console.log('表头行:', headerRow);
                    
                    // 寻找包含"生字"的列
                    for (let i = 0; i < headerRow.length; i++) {
                        const cellValue = String(headerRow[i]).trim();
                        if (cellValue.includes('生字') || cellValue === 'word' || cellValue === '汉字') {
                            characterColIndex = i;
                        } else if (cellValue.includes('拼音') || cellValue === 'pinyin') {
                            pinyinColIndex = i;
                        } else if (cellValue.includes('组词') || cellValue === 'words') {
                            wordsColIndex = i;
                        }
                    }
                    
                    console.log('找到的列索引:', { characterColIndex, pinyinColIndex, wordsColIndex });
                }
                
                currentCopybookData = [];
                
                // 从第二行开始处理数据（跳过表头）
                for (let i = 1; i < jsonData.length; i++) {
                    const row = jsonData[i];
                    if (!row || row.length === 0) continue;
                    
                    let character = '';
                    
                    // 尝试从找到的列获取生字
                    if (characterColIndex !== -1 && row[characterColIndex]) {
                        character = String(row[characterColIndex]).trim();
                    } else {
                        // 如果没有找到特定列，尝试第一列
                        character = String(row[0] || '').trim();
                    }
                    
                    // 只处理非空的中文字符
                    if (character && /[\u4e00-\u9fa5]/.test(character)) {
                        // 获取拼音
                        let pinyin = '';
                        if (pinyinColIndex !== -1 && row[pinyinColIndex]) {
                            pinyin = String(row[pinyinColIndex]).trim();
                        }
                        
                        // 获取组词
                        let words = [];
                        if (wordsColIndex !== -1 && row[wordsColIndex]) {
                            const wordsStr = String(row[wordsColIndex]).trim();
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
                }
                
                console.log('处理完成，找到生字:', currentCopybookData.length);
                
                if (currentCopybookData.length > 0) {
                    currentSource = 'excel';
                    currentGrade = '自定义字帖';
                    generateCopybook('自定义字帖 - 生字练习');
                } else {
                    loadingOverlay.classList.remove('show');
                    alert('Excel文件中未找到有效的生字数据。请确保：\n1. 文件包含"生字"列\n2. 生字列包含中文字符\n3. 文件不是空的');
                }
                
            } catch (error) {
                console.error('处理Excel文件失败:', error);
                loadingOverlay.classList.remove('show');
                alert(`处理Excel文件失败: ${error.message}\n请确保文件格式正确（.xlsx或.xls）`);
            }
        };
        
        reader.onerror = function(error) {
            console.error('读取文件失败:', error);
            loadingOverlay.classList.remove('show');
            alert('读取文件失败，请重试');
        };
        
        // 根据文件类型使用正确的读取方式
        if (file.name.endsWith('.xlsx')) {
            reader.readAsArrayBuffer(file);
        } else {
            reader.readAsBinaryString(file);
        }
    }
    
    // 从系统数据生成字帖
    function generateCopybookFromSystem() {
        if (currentCopybookData.length === 0) return;
        
        currentSource = 'system';
        currentGrade = gradeSelect.value;
        
        const title = `${currentGrade} - 生字练习字帖`;
        generateCopybook(title);
    }
    
    // 生成字帖
    function generateCopybook(title) {
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
                
            } catch (error) {
                console.error('生成字帖失败:', error);
                alert('生成字帖时出错，请重试');
            } finally {
                // 确保隐藏加载提示
                setTimeout(() => {
                    loadingOverlay.classList.remove('show');
                }, 100);
            }
        }, 50);
    }
    
    // 生成字帖内容
    function generateCopybookContent() {
        if (currentCopybookData.length === 0) return;
        
        try {
            // 清空内容
            copybookContent.innerHTML = '';
            
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
                
                // 添加每个生字的字帖单元格
                pageCharacters.forEach((item, index) => {
                    // 第一列显示生字，其他列为空白练习格
                    const isFirstColumn = (index % settings.charsPerRow) === 0;
                    
                    const characterCell = document.createElement('div');
                    characterCell.className = 'character-cell';
                    characterCell.style.width = `${settings.gridSize}px`;
                    characterCell.style.height = `${settings.gridSize}px`;
                    
                    // 添加网格背景
                    const gridBackground = document.createElement('div');
                    gridBackground.className = 'grid-background';
                    gridBackground.classList.add(`grid-${settings.gridType}`);
                    characterCell.appendChild(gridBackground);
                    
                    if (isFirstColumn) {
                        // 第一列显示生字内容
                        const characterContent = document.createElement('div');
                        characterContent.className = 'character-content';
                        characterContent.style.position = 'relative';
                        characterContent.style.zIndex = '1';
                        
                        // 添加生字
                        const characterElement = document.createElement('div');
                        characterElement.className = 'copybook-char';
                        characterElement.textContent = item.character;
                        characterElement.style.fontFamily = settings.fontFamily;
                        characterElement.style.fontSize = `${settings.fontSize}px`;
                        characterElement.style.fontWeight = settings.fontWeight;
                        characterElement.style.color = '#333';
                        characterElement.style.lineHeight = '1';
                        characterContent.appendChild(characterElement);
                        
                        // 添加拼音（如果需要）
                        if (settings.showPinyin && item.pinyin) {
                            const pinyinElement = document.createElement('div');
                            pinyinElement.className = 'copybook-pinyin';
                            pinyinElement.textContent = item.pinyin;
                            pinyinElement.style.fontSize = `${settings.pinyinSize}px`;
                            characterContent.appendChild(pinyinElement);
                        }
                        
                        // 添加笔画数（如果需要）
                        if (settings.showStrokes && item.strokeCount) {
                            const strokesElement = document.createElement('div');
                            strokesElement.className = 'copybook-strokes';
                            strokesElement.textContent = `笔画: ${item.strokeCount}`;
                            strokesElement.style.fontSize = '10px';
                            characterContent.appendChild(strokesElement);
                        }
                        
                        characterCell.appendChild(characterContent);
                    } else {
                        // 其他列为空白练习格
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
            
            // 应用纸张方向
            if (settings.orientation === 'landscape') {
                copybookContent.style.width = '100%';
                copybookContent.style.maxWidth = 'none';
            }
            
        } catch (error) {
            console.error('生成字帖内容失败:', error);
            throw error;
        }
    }
    
    // 估算笔画数（简单版本）
    function estimateStrokeCount(character) {
        // 常见汉字笔画数映射（简化版）
        const strokeMap = {
            '一': 1, '二': 2, '三': 3, '四': 5, '五': 4,
            '六': 4, '七': 2, '八': 2, '九': 2, '十': 2,
            '天': 4, '地': 6, '人': 2, '你': 7, '我': 7,
            '他': 5, '春': 9, '夏': 10, '秋': 9, '冬': 5,
            '日': 4, '月': 4, '水': 4, '火': 4, '土': 3,
            '木': 4, '金': 8, '山': 3, '石': 5, '田': 5,
            '上': 3, '下': 3, '中': 4, '大': 3, '小': 3
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
            
            const pageWidth = settings.orientation === 'portrait' ? 210 : 297;
            const pageHeight = settings.orientation === 'portrait' ? 297 : 210;
            
            // 计算图片尺寸
            const imgWidth = pageWidth - 20; // 左右各10mm边距
            const imgHeight = (canvas.height * imgWidth) / canvas.width;
            
            // 添加图片
            pdf.addImage(imgData, 'PNG', 10, 10, imgWidth, imgHeight);
            
            // 如果图片高度超过页面，需要分页
            let currentY = 10 + imgHeight;
            
            if (currentY > pageHeight - 20) {
                pdf.addPage();
                currentY = 10;
            }
            
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
            pdf.save(`${previewTitle.textContent.replace(/[<>:"/\\|?*]/g, '_')}.pdf`);
            
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
            const content = document.getElementById('copybook-content');
            const clonedContent = content.cloneNode(true);
            
            // 移除不必要的样式
            clonedContent.querySelectorAll('*').forEach(el => {
                el.removeAttribute('style');
            });
            
            // 重新应用基本样式
            clonedContent.querySelectorAll('.character-cell').forEach(cell => {
                cell.style.width = `${settings.gridSize}px`;
                cell.style.height = `${settings.gridSize}px`;
                cell.style.border = '1px solid #ccc';
                cell.style.display = 'flex';
                cell.style.alignItems = 'center';
                cell.style.justifyContent = 'center';
                cell.style.position = 'relative';
            });
            
            // 创建HTML内容
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
                        .copybook-page { 
                            margin-bottom: 20px;
                            page-break-after: always;
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
                            display: flex;
                            align-items: center;
                            justify-content: center;
                            position: relative;
                            width: ${settings.gridSize}px;
                            height: ${settings.gridSize}px;
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
                            z-index: 1; 
                            text-align: center; 
                        }
                        .copybook-char { 
                            font-size: ${settings.fontSize}pt; 
                            margin-bottom: 5px; 
                            font-family: ${settings.fontFamily};
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
                        @media print {
                            .copybook-page { page-break-after: always; }
                            body { margin: 0; }
                        }
                    </style>
                </head>
                <body>
                    ${clonedContent.innerHTML}
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