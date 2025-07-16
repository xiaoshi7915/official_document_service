#!/bin/bash

# 公文写作API服务部署脚本
# 支持源码部署和Docker部署

set -e

echo "========================================="
echo "公文写作API服务部署脚本"
echo "========================================="

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# 打印带颜色的信息
print_info() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# 检查命令是否存在
check_command() {
    if ! command -v $1 &> /dev/null; then
        print_error "$1 命令未找到，请先安装 $1"
        exit 1
    fi
}

# 源码部署
deploy_source() {
    print_info "开始源码部署..."
    
    # 检查Python版本
    check_command python3
    
    # 创建虚拟环境
    if [ ! -d "venv" ]; then
        print_info "创建Python虚拟环境..."
        python3 -m venv venv
    fi
    
    # 激活虚拟环境
    print_info "激活虚拟环境..."
    source venv/bin/activate
    
    # 安装依赖
    print_info "安装Python依赖..."
    pip install -r requirements.txt
    
    # 配置环境变量
    if [ ! -f ".env" ]; then
        if [ -f "env.example" ]; then
            print_warning "未找到.env文件，从env.example复制..."
            cp env.example .env
            print_warning "请编辑.env文件配置OSS信息"
        else
            print_error "未找到env.example文件"
            exit 1
        fi
    fi
    
    print_info "源码部署完成！"
    print_info "启动命令: python run.py"
    print_info "服务地址: http://localhost:8080"
    print_info "API文档: http://localhost:8080/docs"
}

# Docker部署
deploy_docker() {
    print_info "开始Docker部署..."
    
    # 检查Docker
    check_command docker
    
    # 检查docker-compose配置
    if [ ! -f "docker-compose.yml" ]; then
        print_error "未找到docker-compose.yml文件"
        exit 1
    fi
    
    # 构建镜像
    print_info "构建Docker镜像..."
    docker build -t official-writer .
    
    # 启动服务
    print_info "启动Docker容器..."
    docker-compose up -d
    
    print_info "Docker部署完成！"
    print_info "查看日志: docker-compose logs -f"
    print_info "服务地址: http://localhost:8080"
    print_info "停止服务: docker-compose down"
}

# 检查服务状态
check_service() {
    print_info "检查服务状态..."
    
    # 等待服务启动
    sleep 3
    
    if curl -s http://localhost:8080/health > /dev/null; then
        print_info "✅ 服务运行正常"
        print_info "健康检查: curl http://localhost:8080/health"
    else
        print_warning "⚠️ 服务可能未正常启动，请检查日志"
    fi
}

# 主菜单
main_menu() {
    echo ""
    echo "请选择部署方式："
    echo "1) 源码部署 (推荐用于开发)"
    echo "2) Docker部署 (推荐用于生产)"
    echo "3) 仅检查服务状态"
    echo "4) 退出"
    echo ""
    
    read -p "请输入选项 (1-4): " choice
    
    case $choice in
        1)
            deploy_source
            ;;
        2)
            deploy_docker
            ;;
        3)
            check_service
            ;;
        4)
            print_info "退出部署脚本"
            exit 0
            ;;
        *)
            print_error "无效选项，请重新选择"
            main_menu
            ;;
    esac
}

# 检查操作系统
check_os() {
    if [[ "$OSTYPE" == "msys" || "$OSTYPE" == "cygwin" ]]; then
        print_warning "检测到Windows系统，建议使用WSL或Git Bash运行此脚本"
    fi
}

# 主函数
main() {
    check_os
    main_menu
    
    # 部署完成后检查服务
    if [ "$choice" != "3" ] && [ "$choice" != "4" ]; then
        check_service
    fi
    
    echo ""
    print_info "部署完成！"
    echo ""
    echo "📖 更多信息："
    echo "   - 项目文档: README.md"
    echo "   - 快速指南: 快速启动指南.md"
    echo "   - 测试API: python test_api.py"
    echo ""
}

# 脚本入口
main "$@" 