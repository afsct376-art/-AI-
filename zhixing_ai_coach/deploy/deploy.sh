#!/bin/bash
# 知行AI教练 · 一键部署脚本（在服务器上执行）
set -e

echo "=== 知行AI教练 部署脚本 ==="

# 1. 安装 Docker
if ! command -v docker &> /dev/null; then
    echo ">> 安装 Docker..."
    curl -fsSL https://get.docker.com | sh
    systemctl enable docker
    systemctl start docker
fi

# 2. 安装 Docker Compose
if ! command -v docker-compose &> /dev/null; then
    echo ">> 安装 Docker Compose..."
    curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
    chmod +x /usr/local/bin/docker-compose
fi

# 3. 安装 Nginx + SSL（用 acme.sh 免费证书）
if ! command -v nginx &> /dev/null; then
    echo ">> 安装 Nginx..."
    apt-get update && apt-get install -y nginx
fi

# 4. 创建数据目录
mkdir -p data

# 5. 启动服务
echo ">> 启动后端服务..."
docker-compose up -d

# 6. 配置 Nginx（需手动设置域名和SSL）
echo ""
echo "=== ✅ 后端已启动！==="
echo "接下来需要："
echo "1. 把你的域名解析到本服务器IP"
echo "2. 修改 /etc/nginx/sites-enabled/default（用 deploy/nginx.conf）"
echo "3. 申请 SSL 证书：curl https://get.acme.sh | sh && acme.sh --issue -d your-domain.com --nginx"
echo "4. 重启 nginx：systemctl restart nginx"
echo ""
echo "健康检查: curl http://localhost:8000/api/health"
