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
    const printBtnMain = document.getElementById('print-btn-main');
    const applySettingsBtn = document.getElementById('apply-settings');
    const loading = document.getElementById('loading');
    
    // 导出相关元素
    const exportBtn = document.getElementById('export-btn');
    const exportModal = document.getElementById('export-modal');
    const closeExportModal = document.getElementById('close-export-modal');
    const cancelExport = document.getElementById('cancel-export');
    const exportPdf = document.getElementById('export-pdf');
    const exportWord = document.getElementById('export-word');
    const confirmExport = document.getElementById('confirm-export');
    
    // 分页控制元素
    const paginationControls = document.getElementById('pagination-controls');
    const prevPageBtn = document.getElementById('prev-page');
    const nextPageBtn = document.getElementById('next-page');
    const pageInfo = document.getElementById('page-info');
    
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
    
    // 当前字帖数据
    let currentCopybookData = [];
    let currentSource = ''; // 'system' 或 'excel'
    let currentGrade = '';
    
    // 分页相关变量
    let currentPage = 1;
    let totalPages = 1;
    let charsPerPage = 24;
    
    // 导出相关变量
    let selectedExportType = '';
    
    // 当前设置
    let currentSettings = {
        gridSize: 150,
        gridGap: 20,
        gridColor: '#cccccc',
        gridType: 'square',
        fontFamily: "'Ma Shan Zheng', cursive",
        fontSize: 48,
        fontWeight: 400,
        fontColor: '#000000'
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
        
        // 顶部打印按钮
        printBtnMain.addEventListener('click', function() {
            window.print();
        });
        
        // 应用设置按钮
        applySettingsBtn.addEventListener('click', applySettings);
        
        // 分页按钮事件
        prevPageBtn.addEventListener('click', goToPrevPage);
        nextPageBtn.addEventListener('click', goToNextPage);
        
        // 导出相关事件
        exportBtn.addEventListener('click', openExportModal);
        closeExportModal.addEventListener('click', closeExportModalFunc);
        cancelExport.addEventListener('click', closeExportModalFunc);
        
        exportPdf.addEventListener('click', function() {
            selectExportType('pdf');
        });
        
        exportWord.addEventListener('click', function() {
            selectExportType('word');
        });
        
        confirmExport.addEventListener('click', handleExport);
        
        // 设置滑块和选择器事件
        setupSettingsEvents();
        
        // 加载年级数据
        loadAvailableGrades();
        
        // 点击模态框外部关闭
        exportModal.addEventListener('click', function(e) {
            if (e.target === exportModal) {
                closeExportModalFunc();
            }
        });
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
            data.forEach(item => {
                const option = document.createElement('option');
                option.value = item.grade;
                option.textContent = item.grade;
                gradeSelect.appendChild(option);
            });
            
        } catch (error) {
            console.error('加载年级数据失败:', error);
            // 添加默认选项
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
                    pinyin: item.pinyin || '',
                    words: item.words || []
                }));
                
                // 显示字数
                countNumber.textContent = currentCopybookData.length;
                characterCount.style.display = 'block';
                
                console.log(`已加载 ${grade} 的 ${currentCopybookData.length} 个生字`);
            } else {
                // 如果没有找到数据，使用示例数据
                currentCopybookData = getSampleData(grade);
                countNumber.textContent = currentCopybookData.length;
                characterCount.style.display = 'block';
                console.log(`使用示例数据: ${grade} 的 ${currentCopybookData.length} 个生字`);
            }
        } catch (error) {
            console.error('加载年级数据失败:', error);
            // 使用示例数据
            currentCopybookData = getSampleData(grade);
            countNumber.textContent = currentCopybookData.length;
            characterCount.style.display = 'block';
            console.log(`使用示例数据: ${grade} 的 ${currentCopybookData.length} 个生字`);
        }
    }
    
    // 获取示例数据
    function getSampleData(grade) {
        const sampleData = {
            '一年级上册': [
                { character: '一', pinyin: 'yī', words: ['一个', '一起'] },
                { character: '二', pinyin: 'èr', words: ['二月', '第二'] },
                { character: '三', pinyin: 'sān', words: ['三天', '三月'] },
                { character: '四', pinyin: 'sì', words: ['四月', '四周'] },
                { character: '五', pinyin: 'wǔ', words: ['五月', '五个'] },
                { character: '六', pinyin: 'liù', words: ['六月', '六天'] },
                { character: '七', pinyin: 'qī', words: ['七月', '七天'] },
                { character: '八', pinyin: 'bā', words: ['八月', '八天'] },
                { character: '九', pinyin: 'jiǔ', words: ['九月', '九天'] },
                { character: '十', pinyin: 'shí', words: ['十月', '十天'] }
            ],
            '一年级下册': [
                { character: '人', pinyin: 'rén', words: ['人民', '大人'] },
                { character: '口', pinyin: 'kǒu', words: ['口水', '出口'] },
                { character: '手', pinyin: 'shǒu', words: ['手机', '小手'] },
                { character: '足', pinyin: 'zú', words: ['足球', '足够'] },
                { character: '日', pinyin: 'rì', words: ['日期', '生日'] },
                { character: '月', pinyin: 'yuè', words: ['月亮', '月份'] },
                { character: '水', pinyin: 'shuǐ', words: ['水果', '喝水'] },
                { character: '火', pinyin: 'huǒ', words: ['火车', '着火'] },
                { character: '山', pinyin: 'shān', words: ['山上', '爬山'] },
                { character: '石', pinyin: 'shí', words: ['石头', '石子'] }
            ]
        };
        
        return sampleData[grade] || sampleData['一年级上册'];
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
        currentGrade = gradeSelect.value;
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
        // 显示加载提示
        loading.classList.add('show');
        copybookPreview.classList.remove('show');
        
        // 清空预览区
        copybookGrid.innerHTML = '';
        
        // 设置标题
        if (currentSource === 'system') {
            previewTitle.textContent = `${gradeSelect.options[gradeSelect.selectedIndex].text} 生字字帖`;
        } else {
            previewTitle.textContent = `自定义字帖 (${currentCopybookData.length}个汉字)`;
        }
        
        // 根据方格大小动态计算每页字数
        calculateCharsPerPage();
        
        // 计算分页
        calculatePagination();
        
        // 显示预览区
        setTimeout(() => {
            loading.classList.remove('show');
            copybookPreview.classList.add('show');
            
            // 显示分页控制（如果有多页）
            if (totalPages > 1) {
                paginationControls.style.display = 'flex';
            } else {
                paginationControls.style.display = 'none';
            }
            
            // 创建页面容器
            const pageContainer = document.createElement('div');
            pageContainer.className = 'page-container';
            
            // 计算当前页的数据范围
            const startIndex = (currentPage - 1) * charsPerPage;
            const endIndex = Math.min(startIndex + charsPerPage, currentCopybookData.length);
            
            // 创建字帖行
            for (let i = startIndex; i < endIndex; i++) {
                createCopybookRow(pageContainer, currentCopybookData[i], i);
            }
            
            copybookGrid.appendChild(pageContainer);
            
            // 更新分页信息
            updatePaginationInfo();
            
            // 更新CSS变量
            updateCSSVariables();
            
            // 应用行间距
            document.documentElement.style.setProperty('--row-gap', currentSettings.gridGap + 'px');
            
        }, 100);
    }
    
    // 创建字帖行（每行第一个格有字，后面的空格子）
    function createCopybookRow(container, item, index) {
        const gridRow = document.createElement('div');
        gridRow.className = 'grid-row';
        
        // 设置行间距
        gridRow.style.marginBottom = currentSettings.gridGap + 'px';
        
        // 每行显示4个格子（第一个有字，后面3个空白）
        const charsPerRow = 4;
        
        for (let j = 0; j < charsPerRow; j++) {
            const gridItem = document.createElement('div');
            gridItem.className = 'grid-item';
            
            // 设置网格大小
            gridItem.style.width = `${currentSettings.gridSize}px`;
            gridItem.style.height = `${currentSettings.gridSize}px`;
            
            // 创建网格内部结构
            const gridInner = document.createElement('div');
            gridInner.className = `grid-inner ${currentSettings.gridType}`;
            
            // 设置网格线颜色
            gridInner.style.borderColor = currentSettings.gridColor;
            gridInner.style.setProperty('--grid-color', currentSettings.gridColor);
            
            // 特殊网格类型的额外元素
            if (currentSettings.gridType === 'nine') {
                const verticalLine1 = document.createElement('div');
                verticalLine1.className = 'vertical-line-1';
                const verticalLine2 = document.createElement('div');
                verticalLine2.className = 'vertical-line-2';
                gridInner.appendChild(verticalLine1);
                gridInner.appendChild(verticalLine2);
            } else if (currentSettings.gridType === 'rice') {
                const diagonal1 = document.createElement('div');
                diagonal1.className = 'diagonal-1';
                const diagonal2 = document.createElement('div');
                diagonal2.className = 'diagonal-2';
                gridInner.appendChild(diagonal1);
                gridInner.appendChild(diagonal2);
            }
            
            // 只有第一个格子添加汉字
            if (j === 0) {
                // 添加汉字
                const characterElement = document.createElement('div');
                characterElement.className = 'character';
                characterElement.textContent = item.character;
                characterElement.style.fontFamily = currentSettings.fontFamily;
                characterElement.style.fontSize = `${currentSettings.fontSize}px`;
                characterElement.style.fontWeight = currentSettings.fontWeight;
                characterElement.style.color = currentSettings.fontColor;
                
                gridInner.appendChild(characterElement);
                
                // 添加拼音
                if (item.pinyin) {
                    const pinyinElement = document.createElement('div');
                    pinyinElement.className = 'pinyin';
                    pinyinElement.textContent = item.pinyin;
                    gridInner.appendChild(pinyinElement);
                }
                
                // 添加序号
                const orderElement = document.createElement('div');
                orderElement.className = 'stroke-order';
                orderElement.textContent = `${index + 1}`;
                gridInner.appendChild(orderElement);
                
                // 添加组词（如果有）
                if (item.words && item.words.length > 0) {
                    const wordsElement = document.createElement('div');
                    wordsElement.className = 'words';
                    wordsElement.textContent = item.words.slice(0, 2).join(' ');
                    gridInner.appendChild(wordsElement);
                }
            }
            
            gridItem.appendChild(gridInner);
            gridRow.appendChild(gridItem);
        }
        
        container.appendChild(gridRow);
    }
    
    // 根据方格大小计算每页字数
    function calculateCharsPerPage() {
        // 根据方格大小动态调整每页字数
        // 基本规则：方格越大，每页字数越少
        const gridSize = currentSettings.gridSize;
        
        if (gridSize <= 100) {
            charsPerPage = 48; // 小方格，显示更多字
        } else if (gridSize <= 150) {
            charsPerPage = 36;
        } else if (gridSize <= 200) {
            charsPerPage = 24;
        } else if (gridSize <= 250) {
            charsPerPage = 16;
        } else {
            charsPerPage = 12; // 大方格，显示较少字
        }
        
        // 更新选择框
        document.getElementById('chars-per-page').value = charsPerPage;
    }
    
    // 计算分页
    function calculatePagination() {
        totalPages = Math.ceil(currentCopybookData.length / charsPerPage);
        if (currentPage > totalPages) {
            currentPage = totalPages || 1;
        }
    }
    
    // 更新分页信息
    function updatePaginationInfo() {
        pageInfo.textContent = `第 ${currentPage} 页 / 共 ${totalPages} 页`;
        prevPageBtn.disabled = currentPage === 1;
        nextPageBtn.disabled = currentPage === totalPages;
    }
    
    // 上一页
    function goToPrevPage() {
        if (currentPage > 1) {
            currentPage--;
            renderCopybook();
        }
    }
    
    // 下一页
    function goToNextPage() {
        if (currentPage < totalPages) {
            currentPage++;
            renderCopybook();
        }
    }
    
    // 打开导出模态框
    function openExportModal() {
        if (currentCopybookData.length === 0) {
            alert('请先生成字帖再导出');
            return;
        }
        exportModal.classList.add('show');
        selectedExportType = '';
        confirmExport.style.display = 'none';
        
        // 移除所有选项的选中状态
        exportPdf.classList.remove('selected');
        exportWord.classList.remove('selected');
    }
    
    // 关闭导出模态框
    function closeExportModalFunc() {
        exportModal.classList.remove('show');
    }
    
    // 选择导出类型
    function selectExportType(type) {
        selectedExportType = type;
        confirmExport.style.display = 'block';
        
        // 更新选中状态
        exportPdf.classList.remove('selected');
        exportWord.classList.remove('selected');
        
        if (type === 'pdf') {
            exportPdf.classList.add('selected');
            exportPdf.style.borderColor = '#4a6fa5';
            exportPdf.style.backgroundColor = '#f8faff';
            exportWord.style.borderColor = '#e0e0e0';
            exportWord.style.backgroundColor = 'white';
        } else if (type === 'word') {
            exportWord.classList.add('selected');
            exportWord.style.borderColor = '#4a6fa5';
            exportWord.style.backgroundColor = '#f8faff';
            exportPdf.style.borderColor = '#e0e0e0';
            exportPdf.style.backgroundColor = 'white';
        }
    }
    
    // 处理导出
    async function handleExport() {
        if (!selectedExportType) {
            alert('请选择导出格式');
            return;
        }
        
        try {
            // 显示加载提示
            loading.classList.add('show');
            closeExportModalFunc();
            
            if (selectedExportType === 'pdf') {
                await exportAsPDF();
            } else if (selectedExportType === 'word') {
                await exportAsWord();
            }
        } catch (error) {
            console.error('导出失败:', error);
            alert('导出失败，请重试');
        } finally {
            loading.classList.remove('show');
        }
    }
    
    // 导出为PDF
    async function exportAsPDF() {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');
        
        // 设置PDF属性
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        
        // 获取字帖内容
        const copybookContent = document.getElementById('copybook-grid');
        
        // 使用html2canvas将内容转换为图片
        const canvas = await html2canvas(copybookContent, {
            scale: 2, // 提高清晰度
            backgroundColor: '#ffffff',
            useCORS: true,
            logging: false
        });
        
        const imgData = canvas.toDataURL('image/png');
        
        // 计算图片尺寸
        const imgWidth = pageWidth - 20; // 留边距
        const imgHeight = (canvas.height * imgWidth) / canvas.width;
        
        // 添加图片到PDF
        pdf.addImage(imgData, 'PNG', 10, 10, imgWidth, imgHeight);
        
        // 添加标题
        const title = previewTitle.textContent;
        pdf.setFontSize(16);
        pdf.setFont('helvetica', 'bold');
        pdf.text(title, pageWidth / 2, 280, { align: 'center' });
        
        // 添加页脚
        pdf.setFontSize(10);
        pdf.setFont('helvetica', 'normal');
        pdf.text('© 小学语文生字学习系统 - 字帖打印', pageWidth / 2, 287, { align: 'center' });
        
        // 保存文件
        pdf.save(`${title.replace(/\s+/g, '_')}.pdf`);
    }
    
    // 导出为Word
    async function exportAsWord() {
        // 由于浏览器端生成Word比较复杂，这里提供一个简单的替代方案
        // 实际项目中可能需要服务器端支持
        
        alert('Word导出功能需要服务器端支持，当前版本暂不支持Word导出。\n建议使用PDF导出功能。');
        
        // 以下是一个简单的HTML到Word的转换示例
        // const content = document.getElementById('copybook-grid').innerHTML;
        // const blob = new Blob([`
        //     <!DOCTYPE html>
        //     <html>
        //     <head>
        //         <meta charset="UTF-8">
        //         <style>
        //             body { font-family: Arial, sans-serif; margin: 20px; }
        //             .grid-row { display: flex; margin-bottom: ${currentSettings.gridGap}px; }
        //             .grid-item { width: ${currentSettings.gridSize}px; height: ${currentSettings.gridSize}px; border: 1px solid ${currentSettings.gridColor}; margin-right: 10px; display: flex; align-items: center; justify-content: center; }
        //             .character { font-size: ${currentSettings.fontSize}px; font-family: ${currentSettings.fontFamily}; }
        //         </style>
        //     </head>
        //     <body>
        //         <h1>${previewTitle.textContent}</h1>
        //         ${content}
        //     </body>
        //     </html>
        // `], { type: 'application/msword' });
        
        // const link = document.createElement('a');
        // link.href = URL.createObjectURL(blob);
        // link.download = `${previewTitle.textContent.replace(/\s+/g, '_')}.doc`;
        // link.click();
    }
    
    // 添加打印样式
    function addPrintStyles() {
        const printStyles = document.createElement('style');
        printStyles.textContent = `
            @media print {
                body * {
                    visibility: hidden;
                }
                
                .copybook-preview,
                .copybook-preview * {
                    visibility: visible;
                }
                
                .copybook-preview {
                    position: absolute;
                    left: 0;
                    top: 0;
                    width: 100%;
                }
                
                .action-bar,
                .settings-panel,
                .print-controls,
                .navigation-buttons,
                .copybook-options,
                .copybook-header,
                .footer,
                .pagination-controls {
                    display: none !important;
                }
                
                .copybook-grid {
                    box-shadow: none;
                    padding: 0;
                }
                
                .page-container {
                    break-inside: avoid;
                    page-break-inside: avoid;
                    margin-bottom: 20px;
                }
                
                .grid-row {
                    break-inside: avoid;
                    page-break-inside: avoid;
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
        `;
        document.head.appendChild(printStyles);
    }
    
    // 更新CSS变量
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
});