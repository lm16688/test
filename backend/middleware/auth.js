const jwt = require('jsonwebtoken');
const User = require('../models/User');

const auth = async (req, res, next) => {
    try {
        // 从请求头获取token
        const token = req.header('Authorization')?.replace('Bearer ', '');
        
        if (!token) {
            return res.status(401).json({
                success: false,
                message: '请先登录'
            });
        }

        // 验证token
        const decoded = jwt.verify(token, process.env.JWT_SECRET || 'your-secret-key');
        
        // 查找用户
        const user = await User.findOne({ 
            _id: decoded.userId,
            isActive: true 
        });

        if (!user) {
            return res.status(401).json({
                success: false,
                message: '用户不存在或已被禁用'
            });
        }

        // 将用户信息添加到请求对象
        req.user = user;
        req.token = token;
        next();
    } catch (error) {
        console.error('认证错误:', error);
        res.status(401).json({
            success: false,
            message: '认证失败，请重新登录'
        });
    }
};

// 可选认证（不强制要求登录）
const optionalAuth = async (req, res, next) => {
    try {
        const token = req.header('Authorization')?.replace('Bearer ', '');
        
        if (token) {
            const decoded = jwt.verify(token, process.env.JWT_SECRET || 'your-secret-key');
            const user = await User.findOne({ 
                _id: decoded.userId,
                isActive: true 
            });
            
            if (user) {
                req.user = user;
                req.token = token;
            }
        }
        next();
    } catch (error) {
        next();
    }
};

// 管理员认证
const adminAuth = async (req, res, next) => {
    auth(req, res, () => {
        if (req.user.role !== 'admin') {
            return res.status(403).json({
                success: false,
                message: '需要管理员权限'
            });
        }
        next();
    });
};

module.exports = { auth, optionalAuth, adminAuth };
