const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const dotenv = require('dotenv');
const path = require('path');

// 加载环境变量
dotenv.config();

// 导入路由
const authRoutes = require('./routes/auth');
const booklistRoutes = require('./routes/booklists');
const commentRoutes = require('./routes/comments');

// 创建Express应用
const app = express();

// 中间件
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// 静态文件服务（用于部署时）
app.use(express.static(path.join(__dirname, '../frontend')));

// 数据库连接
mongoose.connect(process.env.MONGODB_URI || 'mongodb://localhost:27017/booklist_system', {
    useNewUrlParser: true,
    useUnifiedTopology: true,
})
.then(() => console.log('✅ MongoDB连接成功'))
.catch(err => console.error('❌ MongoDB连接失败:', err));

// 路由
app.use('/api/auth', authRoutes);
app.use('/api/booklists', booklistRoutes);
app.use('/api/comments', commentRoutes);

// 健康检查端点
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'ok', 
        message: '书单分享系统API运行正常',
        timestamp: new Date().toISOString()
    });
});

// 前端路由（单页应用）
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, '../frontend/booklist.html'));
});

// 错误处理中间件
app.use((err, req, res, next) => {
    console.error('❌ 服务器错误:', err.stack);
    res.status(500).json({ 
        success: false, 
        message: '服务器内部错误',
        error: process.env.NODE_ENV === 'development' ? err.message : undefined
    });
});

// 404处理
app.use((req, res) => {
    res.status(404).json({ 
        success: false, 
        message: 'API端点不存在' 
    });
});

// 启动服务器
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 服务器运行在 http://localhost:${PORT}`);
    console.log(`📚 书单分享系统已启动`);
});
