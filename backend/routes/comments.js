const express = require('express');
const router = express.Router();
const { body, param } = require('express-validator');
const commentController = require('../controllers/commentController');
const { auth } = require('../middleware/auth');

// 添加评论
router.post('/booklist/:booklistId', auth, [
    param('booklistId').isMongoId().withMessage('无效的书单ID'),
    body('content').notEmpty().withMessage('评论内容不能为空')
], commentController.addComment);

// 获取书单的评论
router.get('/booklist/:booklistId', [
    param('booklistId').isMongoId().withMessage('无效的书单ID')
], commentController.getComments);

// 更新评论
router.put('/:commentId', auth, [
    param('commentId').isMongoId().withMessage('无效的评论ID'),
    body('content').notEmpty().withMessage('评论内容不能为空')
], commentController.updateComment);

// 删除评论
router.delete('/:commentId', auth, [
    param('commentId').isMongoId().withMessage('无效的评论ID')
], commentController.deleteComment);

// 点赞评论
router.post('/:commentId/like', auth, [
    param('commentId').isMongoId().withMessage('无效的评论ID')
], commentController.likeComment);

module.exports = router;
