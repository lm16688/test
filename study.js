// 生字学习页面JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // 获取DOM元素
    const gradeTitle = document.getElementById('grade-title');
    const characterElement = document.getElementById('character');
    const pinyinElement = document.getElementById('pinyin');
    const wordGroupsElement = document.getElementById('word-groups');
    const sentenceElement = document.getElementById('sentence');
    const speakBtn = document.getElementById('speak-btn');
    const prevBtn = document.getElementById('prev-btn');
    const nextBtn = document.getElementById('next-btn');
    const reviewBtn = document.getElementById('review-btn');
    const masteredBtn = document.getElementById('mastered-btn');
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');
    const totalCountElement = document.getElementById('total-count');
    const learnedCountElement = document.getElementById('learned-count');
    const reviewCountElement = document.getElementById('review-count');
    const masteredCountElement = document.getElementById('mastered-count');
    const hintText = document.getElementById('hint-text');
    
    // 学习状态变量
    let currentGrade = '';
    let charactersData = [];
    let currentIndex = 0;
    let learnedCharacters = [];
    let reviewCharacters = [];
    let masteredCharacters = [];
    
    // 初始化
    async function init() {
        // 从localStorage获取选择的年级
        currentGrade = localStorage.getItem('selectedGrade') || '一年级上册';
        gradeTitle.textContent = currentGrade + ' 生字学习';
        
        // 从JSON文件加载生字数据
        await loadCharactersData();
        
        // 从localStorage加载学习进度
        loadLearningProgress();
        
        // 更新显示
        updateDisplay();
        updateStats();
    }
    
    // 加载生字数据
    async function loadCharactersData() {
        try {
            // 从data.json文件加载数据
            const response = await fetch('data.json');
            if (!response.ok) {
                throw new Error(`HTTP错误! 状态: ${response.status}`);
            }
            
            const data = await response.json();
            
            // 根据新格式查找对应年级的数据
            const gradeData = data.find(item => item.grade === currentGrade);
            
            if (gradeData && gradeData.characters) {
                // 将数据格式转换为内部使用的格式
                charactersData = gradeData.characters.map(item => ({
                    character: item.word,
                    pinyin: item.pinyin,
                    wordGroups: item.words,
                    sentence: item.sentence
                }));
            } else {
                console.warn(`年级"${currentGrade}"没有找到数据，使用演示数据`);
                // 如果没有找到对应年级的数据，使用演示数据作为后备
                charactersData = getDemoData();
            }
            
            // 更新显示
            updateDisplay();
            updateStats();
            
        } catch (error) {
            console.error('加载数据失败:', error);
            // 使用演示数据作为后备
            charactersData = getDemoData();
            updateDisplay();
            updateStats();
        }
    }
    
    // 获取演示数据（作为后备）
    function getDemoData() {
        return [
            {character: "天", pinyin: "tiān", wordGroups: ["天空", "今天"], sentence: "蓝蓝的天空像大海。"},
            {character: "地", pinyin: "dì", wordGroups: ["大地", "土地"], sentence: "大地妈妈真温暖。"},
            {character: "人", pinyin: "rén", wordGroups: ["人们", "好人"], sentence: "人们都在努力工作。"},
            {character: "你", pinyin: "nǐ", wordGroups: ["你好", "你们"], sentence: "你们好，新同学！"},
            {character: "我", pinyin: "wǒ", wordGroups: ["我们", "自我"], sentence: "我们是一年级学生。"}
        ];
    }
    
    // 加载学习进度
    function loadLearningProgress() {
        const savedProgress = localStorage.getItem(`learningProgress_${currentGrade}`);
        if (savedProgress) {
            const progress = JSON.parse(savedProgress);
            learnedCharacters = progress.learned || [];
            reviewCharacters = progress.review || [];
            masteredCharacters = progress.mastered || [];
        } else {
            // 初始状态：所有生字都未学习
            learnedCharacters = [];
            reviewCharacters = [];
            masteredCharacters = [];
        }
    }
    
    // 保存学习进度
    function saveLearningProgress() {
        const progress = {
            learned: learnedCharacters,
            review: reviewCharacters,
            mastered: masteredCharacters,
            lastUpdated: new Date().toISOString()
        };
        localStorage.setItem(`learningProgress_${currentGrade}`, JSON.stringify(progress));
    }
    
    // 更新显示当前生字
    function updateDisplay() {
        if (charactersData.length === 0) {
            characterElement.textContent = "无";
            pinyinElement.textContent = "";
            wordGroupsElement.innerHTML = "";
            sentenceElement.textContent = "暂无数据";
            return;
        }
        
        const currentChar = charactersData[currentIndex];
        
        // 更新生字信息
        characterElement.textContent = currentChar.character;
        pinyinElement.textContent = currentChar.pinyin;
        
        // 更新组词
        wordGroupsElement.innerHTML = "";
        currentChar.wordGroups.forEach(word => {
            const span = document.createElement('span');
            span.className = 'word-group';
            span.textContent = word;
            wordGroupsElement.appendChild(span);
        });
        
        // 更新例句
        sentenceElement.textContent = currentChar.sentence;
        
        // 更新进度条
        const progressPercent = ((currentIndex + 1) / charactersData.length) * 100;
        progressFill.style.width = `${progressPercent}%`;
        progressText.textContent = `${currentIndex + 1}/${charactersData.length}`;
        
        // 更新按钮状态
        prevBtn.disabled = currentIndex === 0;
        nextBtn.disabled = currentIndex === charactersData.length - 1;
        
        // 更新艾宾浩斯提示
        updateHint();
    }
    
    // 更新学习统计
    function updateStats() {
        totalCountElement.textContent = charactersData.length;
        learnedCountElement.textContent = learnedCharacters.length;
        reviewCountElement.textContent = reviewCharacters.length;
        masteredCountElement.textContent = masteredCharacters.length;
    }
    
    // 更新艾宾浩斯提示
    function updateHint() {
        const currentChar = charactersData[currentIndex];
        const char = currentChar.character;
        
        // 检查当前生字的学习状态
        if (masteredCharacters.includes(char)) {
            hintText.textContent = "恭喜！这个生字您已经标记为完全掌握。根据艾宾浩斯记忆法，建议31天后进行一次最终复习以巩固长期记忆。";
        } else if (reviewCharacters.includes(char)) {
            hintText.textContent = "这个生字您标记为需要复习。根据艾宾浩斯记忆法，建议在1小时后、9小时后和1天后分别进行复习。";
        } else if (learnedCharacters.includes(char)) {
            hintText.textContent = "这个生字您已经学习过。根据艾宾浩斯记忆法，建议在20分钟后进行一次复习以巩固记忆。";
        } else {
            hintText.textContent = "这是您第一次学习这个生字。根据艾宾浩斯记忆法，请在20分钟后复习一次，然后按照计划进行后续复习。";
        }
    }
    
    // 语音朗读
    function speakText(text) {
        if ('speechSynthesis' in window) {
            // 停止任何正在进行的朗读
            speechSynthesis.cancel();
            
            // 创建语音实例
            const utterance = new SpeechSynthesisUtterance(text);
            utterance.lang = 'zh-CN';
            utterance.rate = 0.8; // 稍慢的语速
            
            // 开始朗读
            speechSynthesis.speak(utterance);
        } else {
            alert('您的浏览器不支持语音朗读功能，请使用Chrome、Edge或Safari浏览器。');
        }
    }
    
    // 事件监听器
    speakBtn.addEventListener('click', function() {
        const currentChar = charactersData[currentIndex];
        // 朗读生字
        speakText(currentChar.character);
        
        // 添加到已学习列表（如果还未学习）
        if (!learnedCharacters.includes(currentChar.character)) {
            learnedCharacters.push(currentChar.character);
            saveLearningProgress();
            updateStats();
            updateHint();
        }
    });
    
    prevBtn.addEventListener('click', function() {
        if (currentIndex > 0) {
            currentIndex--;
            updateDisplay();
        }
    });
    
    nextBtn.addEventListener('click', function() {
        if (currentIndex < charactersData.length - 1) {
            currentIndex++;
            updateDisplay();
        }
    });
    
    reviewBtn.addEventListener('click', function() {
        const currentChar = charactersData[currentIndex];
        const char = currentChar.character;
        
        // 添加到复习列表
        if (!reviewCharacters.includes(char)) {
            reviewCharacters.push(char);
        }
        
        // 从已掌握列表中移除（如果存在）
        const masteredIndex = masteredCharacters.indexOf(char);
        if (masteredIndex !== -1) {
            masteredCharacters.splice(masteredIndex, 1);
        }
        
        // 保存进度
        saveLearningProgress();
        updateStats();
        updateHint();
        
        // 显示反馈
        alert(`已将"${char}"标记为需要复习。系统会根据艾宾浩斯记忆法提醒您复习。`);
    });
    
    masteredBtn.addEventListener('click', function() {
        const currentChar = charactersData[currentIndex];
        const char = currentChar.character;
        
        // 添加到已掌握列表
        if (!masteredCharacters.includes(char)) {
            masteredCharacters.push(char);
        }
        
        // 从复习列表中移除（如果存在）
        const reviewIndex = reviewCharacters.indexOf(char);
        if (reviewIndex !== -1) {
            reviewCharacters.splice(reviewIndex, 1);
        }
        
        // 确保在已学习列表中
        if (!learnedCharacters.includes(char)) {
            learnedCharacters.push(char);
        }
        
        // 保存进度
        saveLearningProgress();
        updateStats();
        updateHint();
        
        // 显示反馈
        alert(`恭喜！已将"${char}"标记为已掌握。`);
    });
    
    // 初始化应用
    init();
});