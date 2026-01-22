const Booklist = require('../models/Booklist');
const Comment = require('../models/Comment');
const { validationResult } = require('express-validator');

// 获取所有书单
exports.getAllBooklists = async (req, res) => {
    try {
        const { 
            page = 1, 
            limit = 20, 
            subject, 
            year, 
            month, 
            sort = 'newest',
            search,
            creator
        } = req.query;
        
        // 构建查询条件
        const query = { isPublic: true };
        
        if (subject && subject !== 'all') {
            query.subject = subject;
        }
        
        if (year && month) {
            query.year = parseInt(year);
            query.month = parseInt(month);
        } else if (year) {
            query.year = parseInt(year);
        }
        
        if (creator) {
            query.creatorName = { $regex: creator, $options: 'i' };
        }
        
        if (search) {
            query.$or = [
                { title: { $regex: search, $options: 'i' } },
                { content: { $regex: search, $options: 'i' } },
                { tags: { $regex: search, $options: 'i' } }
            ];
        }
        
        // 排序选项
        let sortOption = {};
        switch(sort) {
            case 'newest':
                sortOption = { createdAt: -1 };
                break;
            case 'oldest':
                sortOption = { createdAt: 1 };
                break;
            case 'popular':
                sortOption = { viewCount: -1 };
                break;
            case 'likes':
                sortOption = { likeCount: -1 };
                break;
            case 'comments':
                sortOption = { commentCount: -1 };
                break;
            default:
                sortOption = { createdAt: -1 };
        }
        
        // 分页
        const skip = (parseInt(page) - 1) * parseInt(limit);
        
        // 查询书单
        const booklists = await Booklist.find(query)
            .sort(sortOption)
            .skip(skip)
            .limit(parseInt(limit))
            .populate('creator', 'nickname avatar')
            .lean();
        
        // 获取总数
        const total = await Booklist.countDocuments(query);
        
        // 获取唯一的年月用于筛选
        const yearMonths = await Booklist.aggregate([
            { $match: { isPublic: true } },
            { 
                $group: {
                    _id: { year: '$year', month: '$month' },
                    count: { $sum: 1 }
                }
            },
            { $sort: { '_id.year': -1, '_id.month': -1 } }
        ]);
        
        res.json({
            success: true,
            data: {
                booklists,
                pagination: {
                    page: parseInt(page),
                    limit: parseInt(limit),
                    total,
                    pages: Math.ceil(total / parseInt(limit))
                },
                filters: {
                    yearMonths: yearMonths.map(ym => ({
                        year: ym._id.year,
                        month: ym._id.month,
                        count: ym.count
                    }))
                }
            }
        });
    } catch (error) {
        console.error('获取书单错误:', error);
        res.status(500).json({
            success: false,
            message: '获取书单失败'
        });
    }
};

// 创建书单
exports.createBooklist = async (req, res) => {
    try {
        // 验证输入
        const errors = validationResult(req);
        if (!errors.isEmpty()) {
            return res.status(400).json({
                success: false,
                errors: errors.array()
            });
        }
        
        const { title, content, subject, year, month, bgColor, bgIndex, tags } = req.body;
        
        // 创建书单
        const booklist = new Booklist({
            title,
            content,
            subject,
            year: parseInt(year),
            month: parseInt(month),
            bgColor,
            bgIndex: parseInt(bgIndex),
            creator: req.user._id,
            creatorName: req.user.nickname,
            tags: tags ? tags.split(',').map(tag => tag.trim()) : []
        });
        
        await booklist.save();
        
        res.status(201).json({
            success: true,
            message: '书单创建成功',
            data: {
                booklist: await booklist.populate('creator', 'nickname avatar')
            }
        });
    } catch (error) {
        console.error('创建书单错误:', error);
        res.status(500).json({
            success: false,
            message: '创建书单失败'
        });
    }
};

// 获取单个书单
exports.getBooklist = async (req, res) => {
    try {
        const { id } = req.params;
        
        const booklist = await Booklist.findById(id)
            .populate('creator', 'nickname avatar')
            .populate({
                path: 'comments',
                match: { isDeleted: false },
                populate: {
                    path: 'user',
                    select: 'nickname avatar'
                },
                options: { sort: { createdAt: -1 } }
            });
        
        if (!booklist) {
            return res.status(404).json({
                success: false,
                message: '书单不存在'
            });
        }
        
        // 增加浏览次数
        await booklist.incrementViewCount();
        
        res.json({
            success: true,
            data: { booklist }
        });
    } catch (error) {
        console.error('获取书单详情错误:', error);
        res.status(500).json({
            success: false,
            message: '获取书单详情失败'
        });
    }
};

// 更新书单
exports.updateBooklist = async (req, res) => {
    try {
        const { id } = req.params;
        const updates = req.body;
        
        // 查找书单
        const booklist = await Booklist.findById(id);
        
        if (!booklist) {
            return res.status(404).json({
                success: false,
                message: '书单不存在'
            });
        }
        
        // 检查权限
        if (booklist.creator.toString() !== req.user._id.toString() && req.user.role !== 'admin') {
            return res.status(403).json({
                success: false,
                message: '没有权限修改此书单'
            });
        }
        
        // 更新字段
        if (updates.title) booklist.title = updates.title;
        if (updates.content) booklist.content = updates.content;
        if (updates.subject) booklist.subject = updates.subject;
        if (updates.year) booklist.year = parseInt(updates.year);
        if (updates.month) booklist.month = parseInt(updates.month);
        if (updates.bgColor) booklist.bgColor = updates.bgColor;
        if (updates.bgIndex !== undefined) booklist.bgIndex = parseInt(updates.bgIndex);
        if (updates.tags) booklist.tags = updates.tags.split(',').map(tag => tag.trim());
        
        await booklist.save();
        
        res.json({
            success: true,
            message: '书单更新成功',
            data: {
                booklist: await booklist.populate('creator', 'nickname avatar')
            }
        });
    } catch (error) {
        console.error('更新书单错误:', error);
        res.status(500).json({
            success: false,
            message: '更新书单失败'
        });
    }
};

// 删除书单
exports.deleteBooklist = async (req, res) => {
    try {
        const { id } = req.params;
        
        // 查找书单
        const booklist = await Booklist.findById(id);
        
        if (!booklist) {
            return res.status(404).json({
                success: false,
                message: '书单不存在'
            });
        }
        
        // 检查权限
        if (booklist.creator.toString() !== req.user._id.toString() && req.user.role !== 'admin') {
            return res.status(403).json({
                success: false,
                message: '没有权限删除此书单'
            });
        }
        
        // 删除相关评论
        await Comment.deleteMany({ booklist: id });
        
        // 删除书单
        await booklist.deleteOne();
        
        res.json({
            success: true,
            message: '书单删除成功'
        });
    } catch (error) {
        console.error('删除书单错误:', error);
        res.status(500).json({
            success: false,
            message: '删除书单失败'
        });
    }
};

// 点赞书单
exports.likeBooklist = async (req, res) => {
    try {
        const { id } = req.params;
        
        const booklist = await Booklist.findById(id);
        
        if (!booklist) {
            return res.status(404).json({
                success: false,
                message: '书单不存在'
            });
        }
        
        // 增加点赞数
        await booklist.incrementLikeCount();
        
        res.json({
            success: true,
            message: '点赞成功',
            data: {
                likeCount: booklist.likeCount + 1
            }
        });
    } catch (error) {
        console.error('点赞错误:', error);
        res.status(500).json({
            success: false,
            message: '点赞失败'
        });
    }
};

// 获取用户的书单
exports.getUserBooklists = async (req, res) => {
    try {
        const { userId } = req.params;
        const { page = 1, limit = 20 } = req.query;
        
        const skip = (parseInt(page) - 1) * parseInt(limit);
        
        const booklists = await Booklist.find({ 
            creator: userId,
            isPublic: true 
        })
        .sort({ createdAt: -1 })
        .skip(skip)
        .limit(parseInt(limit))
        .populate('creator', 'nickname avatar');
        
        const total = await Booklist.countDocuments({ 
            creator: userId,
            isPublic: true 
        });
        
        res.json({
            success: true,
            data: {
                booklists,
                pagination: {
                    page: parseInt(page),
                    limit: parseInt(limit),
                    total,
                    pages: Math.ceil(total / parseInt(limit))
                }
            }
        });
    } catch (error) {
        console.error('获取用户书单错误:', error);
        res.status(500).json({
            success: false,
            message: '获取用户书单失败'
        });
    }
};
