const User = require('../models/User');
const jwt = require('jsonwebtoken');
const { validationResult } = require('express-validator');

// 生成JWT Token
const generateToken = (userId) => {
    return jwt.sign(
        { userId },
        process.env.JWT_SECRET || 'your-secret-key',
        { expiresIn: '7d' }
    );
};

// 用户注册
exports.register = async (req, res) => {
    try {
        // 验证输入
        const errors = validationResult(req);
        if (!errors.isEmpty()) {
            return res.status(400).json({
                success: false,
                errors: errors.array()
            });
        }

        const { username, nickname, email, password } = req.body;

        // 检查用户名是否已存在
        const existingUser = await User.findOne({ 
            $or: [
                { username },
                { nickname }
            ]
        });

        if (existingUser) {
            return res.status(400).json({
                success: false,
                message: '用户名或昵称已被使用'
            });
        }

        // 创建新用户
        const user = new User({
            username,
            nickname,
            email,
            password
        });

        await user.save();

        // 生成token
        const token = generateToken(user._id);

        // 更新最后登录时间
        await user.updateLastLogin();

        res.status(201).json({
            success: true,
            message: '注册成功',
            data: {
                user: {
                    id: user._id,
                    username: user.username,
                    nickname: user.nickname,
                    email: user.email,
                    avatar: user.avatar,
                    role: user.role
                },
                token
            }
        });
    } catch (error) {
        console.error('注册错误:', error);
        res.status(500).json({
            success: false,
            message: '注册失败，请稍后重试'
        });
    }
};

// 用户登录
exports.login = async (req, res) => {
    try {
        const { username, password } = req.body;

        // 查找用户
        const user = await User.findOne({ 
            $or: [
                { username },
                { email: username },
                { nickname: username }
            ],
            isActive: true 
        });

        if (!user) {
            return res.status(401).json({
                success: false,
                message: '用户名或密码错误'
            });
        }

        // 验证密码
        const isPasswordValid = await user.comparePassword(password);
        if (!isPasswordValid) {
            return res.status(401).json({
                success: false,
                message: '用户名或密码错误'
            });
        }

        // 生成token
        const token = generateToken(user._id);

        // 更新最后登录时间
        await user.updateLastLogin();

        res.json({
            success: true,
            message: '登录成功',
            data: {
                user: {
                    id: user._id,
                    username: user.username,
                    nickname: user.nickname,
                    email: user.email,
                    avatar: user.avatar,
                    role: user.role
                },
                token
            }
        });
    } catch (error) {
        console.error('登录错误:', error);
        res.status(500).json({
            success: false,
            message: '登录失败，请稍后重试'
        });
    }
};

// 获取当前用户信息
exports.getProfile = async (req, res) => {
    try {
        res.json({
            success: true,
            data: {
                user: {
                    id: req.user._id,
                    username: req.user.username,
                    nickname: req.user.nickname,
                    email: req.user.email,
                    avatar: req.user.avatar,
                    role: req.user.role,
                    createdAt: req.user.createdAt
                }
            }
        });
    } catch (error) {
        console.error('获取用户信息错误:', error);
        res.status(500).json({
            success: false,
            message: '获取用户信息失败'
        });
    }
};

// 更新用户信息
exports.updateProfile = async (req, res) => {
    try {
        const updates = {};
        
        if (req.body.nickname) {
            // 检查昵称是否已存在
            const existingUser = await User.findOne({ 
                nickname: req.body.nickname,
                _id: { $ne: req.user._id }
            });
            
            if (existingUser) {
                return res.status(400).json({
                    success: false,
                    message: '该昵称已被使用'
                });
            }
            updates.nickname = req.body.nickname;
        }
        
        if (req.body.email) {
            updates.email = req.body.email;
        }
        
        if (req.body.avatar) {
            updates.avatar = req.body.avatar;
        }
        
        const user = await User.findByIdAndUpdate(
            req.user._id,
            updates,
            { new: true, runValidators: true }
        );
        
        res.json({
            success: true,
            message: '资料更新成功',
            data: {
                user: {
                    id: user._id,
                    username: user.username,
                    nickname: user.nickname,
                    email: user.email,
                    avatar: user.avatar,
                    role: user.role
                }
            }
        });
    } catch (error) {
        console.error('更新用户信息错误:', error);
        res.status(500).json({
            success: false,
            message: '更新资料失败'
        });
    }
};

// 检查昵称是否可用
exports.checkNickname = async (req, res) => {
    try {
        const { nickname } = req.query;
        
        if (!nickname) {
            return res.status(400).json({
                success: false,
                message: '请提供昵称'
            });
        }
        
        const existingUser = await User.findOne({ nickname });
        
        res.json({
            success: true,
            data: {
                available: !existingUser,
                nickname
            }
        });
    } catch (error) {
        console.error('检查昵称错误:', error);
        res.status(500).json({
            success: false,
            message: '检查昵称失败'
        });
    }
};
