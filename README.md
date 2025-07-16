# 党政机关公文生成API服务

基于GB/T9704-2012标准的党政机关公文自动生成系统，支持Markdown格式输入，自动生成符合规范的Word文档。

## 🚀 主要功能

- ✅ **标准公文格式**：严格按照GB/T9704-2012标准生成公文
- ✅ **Markdown支持**：支持Markdown格式输入，自动转换为规范格式
- ✅ **多级标题**：支持一、二、三级标题的自动编号和格式化
- ✅ **附件支持**：支持最多3个附件，包含表格、文本等多种格式
- ✅ **自动上传**：生成后自动上传到阿里云OSS，支持直接下载
- ✅ **字体规范**：自动应用标准字体（仿宋_GB2312、黑体等）
- ✅ **版式标准**：自动设置标准页边距、行距、字号

## 📋 系统要求

- Python 3.8+
- FastAPI
- python-docx
- oss2
- markdown

## 🛠 快速安装

### 方法1：Docker方式（推荐）

```bash
# 1. 克隆代码
git clone <repository-url>
cd official_writer

# 2. 配置环境变量
cp env.example .env
# 编辑.env文件，填入OSS配置信息

# 3. 使用Docker Compose启动
docker-compose up -d

# 4. 访问服务
curl http://localhost:8080/health
```

### 方法2：本地安装

```bash
# 1. 克隆代码
git clone <repository-url>
cd official_writer

# 2. 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Linux/Mac
# 或 venv\Scripts\activate  # Windows

# 3. 安装依赖
pip install -r requirements.txt

# 4. 配置环境变量
/data/official_writer/venv/bin/python3 -m pip install --upgrade pipcp env.example .env
# 编辑.env文件

# 5. 启动服务
python run.py
```

## 🔧 配置说明

在`.env`文件中配置以下参数：

```env
# 基础配置
APP_HOST=0.0.0.0
APP_PORT=8080
DEBUG=false
API_TOKEN=your-secret-token-12345

# 阿里云OSS配置
OSS_ACCESS_KEY_ID=your-access-key-id
OSS_ACCESS_KEY_SECRET=your-access-key-secret
OSS_ENDPOINT=https://oss-cn-shanghai.aliyuncs.com
OSS_BUCKET_NAME=your-bucket-name
```


## 🔍 测试方法

## 手动测试

```bash
curl -X POST "http://localhost:8080/generate_document" \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer your-secret-token-12345" \
  -d '{
    "content": "## 一、基本要求\n\n各部门要严格按照公文写作规范要求，确保公文质量。",
    "title": "关于加强公文写作规范的通知",
    "issuing_department": "办公厅",
    "issue_date": "2024年1月15日",
    "receiving_department": "各部门",
    "has_attachments": false
  }'
```

## 🚀 部署说明

### Docker部署

1. 确保Docker和Docker Compose已安装
2. 配置`.env`文件
3. 运行：`docker-compose up -d`
4. 检查状态：`docker-compose ps`


## 📄 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。 