const express = require('express');
const cors = require('cors');
const path = require('path');

const app = express();

// 启用CORS，允许所有来源
app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// 数据存储
let users = [];
let booklists = [];
let comments = [];
let userCounter = 1;
let booklistCounter = 1;
let commentCounter = 1;

// 简单的内存存储函数
const saveData = () => {
    // 这里可以添加保存到文件的功能
    console.log('数据已更新');
};

// 健康检查
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'ok', 
        message: '书单分享系统API运行正常',
        timestamp: new Date().toISOString()
    });
});

// 用户注册
app.post('/api/auth/register', (req, res) => {
    const { username, nickname, email, password } = req.body;
    
    // 检查用户名是否已存在
    if (users.find(u => u.username === username)) {
        return res.status(400).json({
            success: false,
            message: '用户名已存在'
        });
    }
    
    // 检查昵称是否已存在
    if (users.find(u => u.nickname === nickname)) {
        return res.status(400).json({
            success: false,
            message: '昵称已存在'
        });
    }
    
    const newUser = {
        id: userCounter++,
        username,
        nickname,
        email,
        password, // 注意：实际应用中需要加密
        avatar: 'default-avatar.png',
        role: 'user',
        createdAt: new Date().toISOString()
    };
    
    users.push(newUser);
    saveData();
    
    // 生成token（简化版）
    const token = `token_${newUser.id}_${Date.now()}`;
    
    res.status(201).json({
        success: true,
        message: '注册成功',
        data: {
            user: {
                id: newUser.id,
                username: newUser.username,
                nickname: newUser.nickname,
                email: newUser.email,
                avatar: newUser.avatar,
                role: newUser.role
            },
            token
        }
    });
});

// 用户登录
app.post('/api/auth/login', (req, res) => {
    const { username, password } = req.body;
    
    // 查找用户
    const user = users.find(u => 
        u.username === username || 
        u.nickname === username || 
        u.email === username
    );
    
    if (!user || user.password !== password) {
        return res.status(401).json({
            success: false,
            message: '用户名或密码错误'
        });
    }
    
    // 生成token
    const token = `token_${user.id}_${Date.now()}`;
    
    res.json({
        success: true,
        message: '登录成功',
        data: {
            user: {
                id: user.id,
                username: user.username,
                nickname: user.nickname,
                email: user.email,
                avatar: user.avatar,
                role: user.role
            },
            token
        }
    });
});

// 检查昵称是否可用
app.get('/api/auth/check-nickname', (req, res) => {
    const { nickname } = req.query;
    
    if (!nickname) {
        return res.status(400).json({
            success: false,
            message: '请提供昵称'
        });
    }
    
    const existingUser = users.find(u => u.nickname === nickname);
    
    res.json({
        success: true,
        data: {
            available: !existingUser,
            nickname
        }
    });
});

// 获取当前用户信息
app.get('/api/auth/profile', (req, res) => {
    const token = req.headers.authorization?.replace('Bearer ', '');
    
    if (!token) {
        return res.status(401).json({
            success: false,
            message: '请先登录'
        });
    }
    
    // 解析token获取用户ID
    const userId = parseInt(token.split('_')[1]);
    const user = users.find(u => u.id === userId);
    
    if (!user) {
        return res.status(401).json({
            success: false,
            message: '用户不存在'
        });
    }
    
    res.json({
        success: true,
        data: {
            user: {
                id: user.id,
                username: user.username,
                nickname: user.nickname,
                email: user.email,
                avatar: user.avatar,
                role: user.role,
                createdAt: user.createdAt
            }
        }
    });
});

// 获取所有书单
app.get('/api/booklists', (req, res) => {
    const { subject, search, page = 1, limit = 20 } = req.query;
    
    let filteredBooklists = [...booklists];
    
    // 按科目筛选
    if (subject && subject !== 'all') {
        filteredBooklists = filteredBooklists.filter(b => b.subject === subject);
    }
    
    // 搜索
    if (search) {
        filteredBooklists = filteredBooklists.filter(b => 
            b.title.toLowerCase().includes(search.toLowerCase()) ||
            b.content.toLowerCase().includes(search.toLowerCase())
        );
    }
    
    // 分页
    const startIndex = (page - 1) * limit;
    const endIndex = page * limit;
    const paginatedBooklists = filteredBooklists.slice(startIndex, endIndex);
    
    // 获取用户信息
    const booklistsWithUser = paginatedBooklists.map(booklist => {
        const user = users.find(u => u.id === booklist.creatorId);
        return {
            ...booklist,
            creator: {
                _id: user?.id,
                nickname: user?.nickname,
                avatar: user?.avatar
            }
        };
    });
    
    // 获取年月筛选选项
    const yearMonths = [...new Set(booklists.map(b => `${b.year}-${b.month}`))]
        .map(ym => {
            const [year, month] = ym.split('-');
            const count = booklists.filter(b => 
                b.year === parseInt(year) && b.month === parseInt(month)
            ).length;
            return { year: parseInt(year), month: parseInt(month), count };
        })
        .sort((a, b) => {
            if (a.year !== b.year) return b.year - a.year;
            return b.month - a.month;
        });
    
    res.json({
        success: true,
        data: {
            booklists: booklistsWithUser,
            pagination: {
                page: parseInt(page),
                limit: parseInt(limit),
                total: filteredBooklists.length,
                pages: Math.ceil(filteredBooklists.length / limit)
            },
            filters: {
                yearMonths
            }
        }
    });
});

// 创建书单
app.post('/api/booklists', (req, res) => {
    const token = req.headers.authorization?.replace('Bearer ', '');
    
    if (!token) {
        return res.status(401).json({
            success: false,
            message: '请先登录'
        });
    }
    
    const userId = parseInt(token.split('_')[1]);
    const user = users.find(u => u.id === userId);
    
    if (!user) {
        return res.status(401).json({
            success: false,
            message: '用户不存在'
        });
    }
    
    const { title, content, subject, year, month, bgIndex, tags } = req.body;
    
    const newBooklist = {
        _id: booklistCounter++,
        title,
        content,
        subject,
        year: parseInt(year),
        month: parseInt(month),
        bgIndex: parseInt(bgIndex),
        bgColor: getBgColor(bgIndex),
        creatorId: user.id,
        creatorName: user.nickname,
        viewCount: 0,
        likeCount: 0,
        commentCount: 0,
        tags: tags ? tags.split(',').map(tag => tag.trim()) : [],
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
    };
    
    booklists.unshift(newBooklist);
    saveData();
    
    res.status(201).json({
        success: true,
        message: '书单创建成功',
        data: {
            booklist: {
                ...newBooklist,
                creator: {
                    _id: user.id,
                    nickname: user.nickname,
                    avatar: user.avatar
                }
            }
        }
    });
});

// 获取单个书单
app.get('/api/booklists/:id', (req, res) => {
    const { id } = req.params;
    const booklistId = parseInt(id);
    
    const booklist = booklists.find(b => b._id === booklistId);
    
    if (!booklist) {
        return res.status(404).json({
            success: false,
            message: '书单不存在'
        });
    }
    
    // 增加浏览次数
    booklist.viewCount++;
    
    // 获取用户信息
    const user = users.find(u => u.id === booklist.creatorId);
    
    // 获取评论
    const booklistComments = comments.filter(c => c.booklistId === booklistId && !c.isDeleted);
    
    res.json({
        success: true,
        data: {
            booklist: {
                ...booklist,
                creator: {
                    _id: user?.id,
                    nickname: user?.nickname,
                    avatar: user?.avatar
                },
                comments: booklistComments.map(comment => {
                    const commentUser = users.find(u => u.id === comment.userId);
                    return {
                        ...comment,
                        user: {
                            _id: commentUser?.id,
                            nickname: commentUser?.nickname,
                            avatar: commentUser?.avatar
                        }
                    };
                })
            }
        }
    });
});

// 更新书单
app.put('/api/booklists/:id', (req, res) => {
    const token = req.headers.authorization?.replace('Bearer ', '');
    
    if (!token) {
        return res.status(401).json({
            success: false,
            message: '请先登录'
        });
    }
    
    const userId = parseInt(token.split('_')[1]);
    const { id } = req.params;
    const booklistId = parseInt(id);
    
    const booklist = booklists.find(b => b._id === booklistId);
    
    if (!booklist) {
        return res.status(404).json({
            success: false,
            message: '书单不存在'
        });
    }
    
    if (booklist.creatorId !== userId) {
        return res.status(403).json({
            success: false,
            message: '没有权限修改此书单'
        });
    }
    
    const { title, content, subject, year, month, bgIndex } = req.body;
    
    booklist.title = title || booklist.title;
    booklist.content = content || booklist.content;
    booklist.subject = subject || booklist.subject;
    booklist.year = year ? parseInt(year) : booklist.year;
    booklist.month = month ? parseInt(month) : booklist.month;
    booklist.bgIndex = bgIndex !== undefined ? parseInt(bgIndex) : booklist.bgIndex;
    booklist.bgColor = getBgColor(booklist.bgIndex);
    booklist.updatedAt = new Date().toISOString();
    
    saveData();
    
    const user = users.find(u => u.id === userId);
    
    res.json({
        success: true,
        message: '书单更新成功',
        data: {
            booklist: {
                ...booklist,
                creator: {
                    _id: user?.id,
                    nickname: user?.nickname,
                    avatar: user?.avatar
                }
            }
        }
    });
});

// 删除书单
app.delete('/api/booklists/:id', (req, res) => {
    const token = req.headers.authorization?.replace('Bearer ', '');
    
    if (!token) {
        return res.status(401).json({
            success: false,
            message: '请先登录'
        });
    }
    
    const userId = parseInt(token.split('_')[1]);
    const { id } = req.params;
    const booklistId = parseInt(id);
    
    const booklistIndex = booklists.findIndex(b => b._id === booklistId);
    
    if (booklistIndex === -1) {
        return res.status(404).json({
            success: false,
            message: '书单不存在'
        });
    }
    
    const booklist = booklists[booklistIndex];
    
    if (booklist.creatorId !== userId) {
        return res.status(403).json({
            success: false,
            message: '没有权限删除此书单'
        });
    }
    
    // 删除相关评论
    comments = comments.filter(c => c.booklistId !== booklistId);
    
    // 删除书单
    booklists.splice(booklistIndex, 1);
    saveData();
    
    res.json({
        success: true,
        message: '书单删除成功'
    });
});

// 添加评论
app.post('/api/comments/booklist/:booklistId', (req, res) => {
    const token = req.headers.authorization?.replace('Bearer ', '');
    
    if (!token) {
        return res.status(401).json({
            success: false,
            message: '请先登录'
        });
    }
    
    const userId = parseInt(token.split('_')[1]);
    const user = users.find(u => u.id === userId);
    
    if (!user) {
        return res.status(401).json({
            success: false,
            message: '用户不存在'
        });
    }
    
    const { booklistId } = req.params;
    const booklistIdNum = parseInt(booklistId);
    const { content } = req.body;
    
    const booklist = booklists.find(b => b._id === booklistIdNum);
    
    if (!booklist) {
        return res.status(404).json({
            success: false,
            message: '书单不存在'
        });
    }
    
    const newComment = {
        _id: commentCounter++,
        content,
        booklistId: booklistIdNum,
        userId,
        userName: user.nickname,
        isDeleted: false,
        createdAt: new Date().toISOString()
    };
    
    comments.unshift(newComment);
    
    // 更新书单评论数
    booklist.commentCount = comments.filter(c => c.booklistId === booklistIdNum && !c.isDeleted).length;
    saveData();
    
    res.status(201).json({
        success: true,
        message: '评论添加成功',
        data: {
            comment: {
                ...newComment,
                user: {
                    _id: user.id,
                    nickname: user.nickname,
                    avatar: user.avatar
                }
            }
        }
    });
});

// 删除评论
app.delete('/api/comments/:commentId', (req, res) => {
    const token = req.headers.authorization?.replace('Bearer ', '');
    
    if (!token) {
        return res.status(401).json({
            success: false,
            message: '请先登录'
        });
    }
    
    const userId = parseInt(token.split('_')[1]);
    const { commentId } = req.params;
    const commentIdNum = parseInt(commentId);
    
    const commentIndex = comments.findIndex(c => c._id === commentIdNum);
    
    if (commentIndex === -1) {
        return res.status(404).json({
            success: false,
            message: '评论不存在'
        });
    }
    
    const comment = comments[commentIndex];
    
    if (comment.userId !== userId) {
        return res.status(403).json({
            success: false,
            message: '没有权限删除此评论'
        });
    }
    
    // 软删除
    comment.isDeleted = true;
    comment.content = '[评论已删除]';
    
    // 更新书单评论数
    const booklist = booklists.find(b => b._id === comment.booklistId);
    if (booklist) {
        booklist.commentCount = comments.filter(c => c.booklistId === comment.booklistId && !c.isDeleted).length;
    }
    
    saveData();
    
    res.json({
        success: true,
        message: '评论删除成功'
    });
});

// 辅助函数：获取背景颜色
function getBgColor(index) {
    const colors = [
        '#4a90e2', '#50c878', '#ff7f50', '#9370db', '#ff6b6b',
        'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
        'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)',
        'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)',
        'linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)',
        'linear-gradient(135deg, #fa709a 0%, #fee140 100%)'
    ];
    return colors[index] || colors[0];
}

// 提供静态文件
app.use(express.static(__dirname));

// 默认路由返回前端页面
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// 404处理
app.use((req, res) => {
    res.status(404).json({
        success: false,
        message: 'API端点不存在'
    });
});

// 错误处理
app.use((err, req, res, next) => {
    console.error('服务器错误:', err);
    res.status(500).json({
        success: false,
        message: '服务器内部错误'
    });
});

// 启动服务器
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 服务器运行在 http://localhost:${PORT}`);
    console.log(`📚 书单分享系统已启动`);
    
    // 添加一些测试数据
    if (users.length === 0) {
        users.push({
            id: userCounter++,
            username: 'test',
            nickname: '测试用户',
            email: 'test@example.com',
            password: '123456',
            avatar: 'default-avatar.png',
            role: 'user',
            createdAt: new Date().toISOString()
        });
        
        booklists.push({
            _id: booklistCounter++,
            title: '小学生必读经典书目',
            content: '1.《安徒生童话》\n2.《格林童话》\n3.《小王子》\n4.《西游记》儿童版\n5.《三国演义》儿童版',
            subject: '语文',
            year: 2024,
            month: 1,
            bgIndex: 0,
            bgColor: getBgColor(0),
            creatorId: 1,
            creatorName: '测试用户',
            viewCount: 10,
            likeCount: 5,
            commentCount: 2,
            tags: ['经典', '必读'],
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
        });
        
        comments.push(
            {
                _id: commentCounter++,
                content: '这个书单太好了，正是我需要的！',
                booklistId: 1,
                userId: 1,
                userName: '测试用户',
                isDeleted: false,
                createdAt: new Date(Date.now() - 3600000).toISOString()
            },
            {
                _id: commentCounter++,
                content: '感谢分享，我家孩子很喜欢这些书',
                booklistId: 1,
                userId: 1,
                userName: '测试用户',
                isDeleted: false,
                createdAt: new Date().toISOString()
            }
        );
        
        console.log('✅ 测试数据已添加');
    }
});
