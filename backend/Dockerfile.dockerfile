FROM node:18-alpine

WORKDIR /app

# 复制依赖文件
COPY package*.json ./

# 安装依赖
RUN npm ci --only=production

# 复制源代码
COPY . .

# 创建必要的目录
RUN mkdir -p uploads logs

# 设置权限
RUN chown -R node:node /app

# 切换用户
USER node

# 暴露端口
EXPOSE 3001

# 启动命令
CMD ["node", "server.js"]