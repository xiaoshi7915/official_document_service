"""
公文写作API服务 - FastAPI应用主文件
"""
import logging
from typing import List
from datetime import datetime
from fastapi import FastAPI, HTTPException, Depends, Security, Request, File, UploadFile
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.config import settings
from app.models.request_models import (
    DocumentGenerateRequest, 
    DocumentGenerateWithAttachmentsRequest,
    DocumentGenerateWithoutAttachmentsRequest,
    FileUploadRequest
)
from app.models.response_models import DocumentGenerateResponse, ErrorResponse, FileInfo
from app.services.document_generator import document_generator
from app.services.oss_service import oss_service
from app.services.attachment_processor import attachment_processor
from typing import List, Dict, Any
import json
import re

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 创建FastAPI应用实例
app = FastAPI(
    title="公文写作API服务",
    description="基于GB/T9704-2012标准的党政机关公文生成服务",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json"
)

# 添加CORS中间件
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 在生产环境中应该设置具体的域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# HTTP Bearer认证（可选）
security = HTTPBearer(auto_error=False)

# 全局变量存储上传的附件
uploaded_attachments = {}

def parse_string_array_attachments(attachments_list: List[str]) -> list:
    """
    解析字符串数组格式的附件数据
    
    Args:
        attachments_list: 字符串数组格式的附件数据
        
    Returns:
        list: 解析后的附件列表
    """
    try:
        result = []
        for i, attachment_content in enumerate(attachments_list, 1):
            if not attachment_content or not attachment_content.strip():
                continue
                
            # 智能判断附件类型
            attachment_type = detect_attachment_type(attachment_content)
            
            result.append({
                "order": str(i),
                "type": attachment_type,
                "name": f"附件{i}",
                "markdown_content": attachment_content.strip()
            })
        
        return result
    except Exception as e:
        logger.error(f"解析字符串数组附件失败: {str(e)}")
        return []

def detect_attachment_type(content: str) -> str:
    """
    智能检测附件类型
    
    Args:
        content: 附件内容
        
    Returns:
        str: 附件类型 (table/text/mixed)
    """
    try:
        content = content.strip()
        
        # 检查是否包含表格
        has_table = "|" in content and "---" in content
        
        # 检查是否包含普通文本
        lines = content.split('\n')
        text_lines = [line for line in lines if line.strip() and not line.strip().startswith('|') and '---' not in line]
        has_text = len(text_lines) > 0
        
        if has_table and has_text:
            return "mixed"  # 表格和文本并存
        elif has_table:
            return "table"  # 纯表格
        else:
            return "text"   # 纯文本
            
    except Exception as e:
        logger.error(f"检测附件类型失败: {str(e)}")
        return "text"  # 默认为文本类型

def parse_attachments_string(attachments_str: str) -> list:
    """
    解析字符串格式的附件数据（兼容旧格式）
    
    Args:
        attachments_str: 字符串格式的附件数据
        
    Returns:
        list: 解析后的附件列表
    """
    try:
        # 尝试JSON解析
        if attachments_str.startswith('[') and attachments_str.endswith(']'):
            attachments_list = json.loads(attachments_str)
            result = []
            for i, att in enumerate(attachments_list, 1):
                if isinstance(att, str):
                    # 如果是纯字符串，创建一个默认附件
                    result.append({
                        "order": str(i),
                        "type": detect_attachment_type(att),
                        "name": f"附件{i}",
                        "markdown_content": att
                    })
                elif isinstance(att, dict):
                    result.append(att)
            return result
        else:
            # 如果不是JSON格式，当作单个附件处理
            return [{
                "order": "1",
                "type": detect_attachment_type(attachments_str),
                "name": "附件1",
                "markdown_content": attachments_str
            }]
    except Exception as e:
        logger.error(f"解析附件字符串失败: {str(e)}")
        return []

def verify_token(credentials: HTTPAuthorizationCredentials = Security(security)) -> bool:
    """
    验证API Token
    
    Args:
        credentials: HTTP认证凭据
        
    Returns:
        bool: 验证是否成功
        
    Raises:
        HTTPException: 认证失败时抛出异常
    """
    if not credentials:
        raise HTTPException(
            status_code=401,
            detail="缺少认证Token",
            headers={"WWW-Authenticate": "Bearer"},
        )
    
    if credentials.credentials != settings.API_TOKEN:
        raise HTTPException(
            status_code=401,
            detail="无效的API Token",
            headers={"WWW-Authenticate": "Bearer"},
        )
    return True

@app.get("/")
async def root():
    """根路径 - 服务状态检查"""
    return {
        "service": "公文写作API服务",
        "version": "1.0.0",
        "status": "running",
        "description": "基于GB/T9704-2012标准的党政机关公文生成服务",
        "port": settings.APP_PORT,
        "endpoints": {
            "generate_document": "POST /generate_document",
            "upload_file": "POST /upload_file",
            "generate_document_with_attachments": "POST /generate_document_with_attachments",
            "generate_document_without_attachments": "POST /generate_document_without_attachments",
            "health": "GET /health",
            "docs": "GET /docs",
            "redoc": "GET /redoc",
            "openapi": "GET /openapi.json"
        }
    }

@app.get("/api-info")
async def api_info():
    """API信息"""
    return {
        "title": app.title,
        "description": app.description,
        "version": app.version,
        "docs_url": app.docs_url,
        "redoc_url": app.redoc_url,
        "openapi_url": app.openapi_url
    }

@app.get("/health")
async def health_check(request: Request):
    """健康检查端点"""
    try:
        # 检查OSS连接
        if oss_service is None:
            oss_status = False
            oss_error = "OSS服务未初始化"
        else:
            oss_status = oss_service.check_bucket_exists()
            oss_error = None
        
        # 获取客户端信息
        client_host = request.client.host if request.client else "unknown"
        
        return {
            "status": "healthy",
            "oss_connection": "ok" if oss_status else "error",
            "oss_error": oss_error,
            "oss_config": {
                "endpoint": settings.OSS_ENDPOINT,
                "bucket": settings.OSS_BUCKET_NAME
            },
            "server_config": {
                "host": settings.APP_HOST,
                "port": settings.APP_PORT,
                "debug": settings.DEBUG
            },
            "client_info": {
                "ip": client_host,
                "headers": dict(request.headers)
            },
            "timestamp": "2024-01-15T10:00:00Z"
        }
    except Exception as e:
        logger.error(f"健康检查失败: {str(e)}")
        return JSONResponse(
            status_code=503,
            content={
                "status": "unhealthy",
                "error": str(e),
                "oss_connection": "error",
                "timestamp": "2024-01-15T10:00:00Z"
            }
        )

@app.get("/network-check")
async def network_check(request: Request):
    """网络诊断端点"""
    import socket
    import platform
    
    try:
        # 获取服务器信息
        hostname = socket.gethostname()
        local_ip = socket.gethostbyname(hostname)
        
        return {
            "server_info": {
                "hostname": hostname,
                "local_ip": local_ip,
                "platform": platform.platform(),
                "python_version": platform.python_version(),
                "listening_on": f"{settings.APP_HOST}:{settings.APP_PORT}"
            },
            "client_info": {
                "remote_addr": request.client.host if request.client else "unknown",
                "user_agent": request.headers.get("user-agent", "unknown"),
                "forwarded_for": request.headers.get("x-forwarded-for", "none"),
                "real_ip": request.headers.get("x-real-ip", "none")
            },
            "network_config": {
                "app_host": settings.APP_HOST,
                "app_port": settings.APP_PORT,
                "cors_enabled": True,
                "allowed_origins": "*"
            }
        }
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": f"网络检查失败: {str(e)}"}
        )

@app.post(
    "/generate_document",
    response_model=DocumentGenerateResponse,
    responses={
        400: {"model": ErrorResponse},
        401: {"model": ErrorResponse},
        500: {"model": ErrorResponse}
    }
)
async def generate_official_document(
    request: DocumentGenerateRequest,
    token_valid: bool = Depends(verify_token)
) -> DocumentGenerateResponse:
    """
    生成公文Word文档
    
    Args:
        request: 文档生成请求
        token_valid: Token验证结果
        
    Returns:
        DocumentGenerateResponse: 文档生成响应
        
    Raises:
        HTTPException: 各种错误情况
    """
    try:
        logger.info(f"开始生成公文: {request.title}")
        
        # 验证请求参数 - 支持两种格式
        content = request.markdown_content or request.content
        if not content or not content.strip():
            raise HTTPException(
                status_code=400,
                detail="正文内容不能为空"
            )
        
        if not request.title.strip():
            raise HTTPException(
                status_code=400,
                detail="文档标题不能为空"
            )
        
        if not request.issuing_department.strip():
            raise HTTPException(
                status_code=400,
                detail="发文部门不能为空"
            )
        
        if not request.issue_date.strip():
            raise HTTPException(
                status_code=400,
                detail="发文日期不能为空"
            )
        
        # 验证附件参数
        if request.has_attachments:
            if not request.attachments or len(request.attachments) == 0:
                raise HTTPException(
                    status_code=400,
                    detail="设置为有附件但未提供附件内容"
                )
            
            if len(request.attachments) > 3:
                raise HTTPException(
                    status_code=400,
                    detail="附件数量不能超过3个"
                )
        
        # 准备附件数据 - 支持混合格式
        attachments_data = None
        if request.has_attachments and request.attachments:
            # 验证附件数量
            if len(request.attachments) > 3:
                raise HTTPException(
                    status_code=400,
                    detail="附件数量不能超过3个"
                )
            
            attachments_data = []
            for i, attachment in enumerate(request.attachments, 1):
                if isinstance(attachment, str):
                    # 字符串格式：转换为标准格式
                    if not attachment.strip():
                        raise HTTPException(
                            status_code=400,
                            detail=f"附件{i}内容不能为空"
                        )
                    
                    attachment_type = detect_attachment_type(attachment)
                    attachments_data.append({
                        "order": str(i),
                        "type": attachment_type,
                        "name": f"附件{i}",
                        "markdown_content": attachment.strip()
                    })
                elif isinstance(attachment, dict):
                    # 字典格式：验证必需字段
                    required_fields = ["order", "type", "name", "markdown_content"]
                    for field in required_fields:
                        if field not in attachment:
                            raise HTTPException(
                                status_code=400,
                                detail=f"附件{i}缺少必需字段: {field}"
                            )
                    
                    attachments_data.append({
                        "order": attachment["order"],
                        "type": attachment["type"],
                        "name": attachment["name"],
                        "markdown_content": attachment["markdown_content"]
                    })
                else:
                    # 对象格式：直接使用
                    attachments_data.append({
                        "order": attachment.order,
                        "type": attachment.type,
                        "name": attachment.name,
                        "markdown_content": attachment.markdown_content
                    })
        
        # 生成Word文档
        logger.info("开始生成Word文档")
        document_bytes = document_generator.generate_document(
            title=request.title,
            issuing_department=request.issuing_department,
            issue_date=request.issue_date,
            content=content,
            receiving_department=request.receiving_department,
            has_attachments=request.has_attachments,
            attachments=attachments_data
        )
        
        # 上传到OSS
        logger.info("开始上传文档到OSS")
        
        if oss_service is None:
            logger.error("OSS服务未初始化")
            raise HTTPException(
                status_code=500,
                detail="OSS服务未初始化，无法上传文档"
            )
        
        success, message, download_url = oss_service.upload_document(
            file_content=document_bytes,
            title=request.title,
            issue_date=request.issue_date
        )
        
        if not success:
            logger.error(f"上传文档失败: {message}")
            raise HTTPException(
                status_code=500,
                detail=f"上传文档失败: {message}"
            )
        
        # 生成文件名
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = request.title.replace('/', '_').replace('\\', '_')
        file_name = f"{timestamp}_{safe_title}.docx"
        
        logger.info(f"文档生成成功: {file_name}")
        
        # 返回新格式响应
        return DocumentGenerateResponse(
            body="文档生成成功",
            status_code=200,
            headers={
                "Content-Type": "application/json",
                "X-Generated-At": timestamp,
                "X-Service": "official-document-generator"
            },
            files=[{
                "download_url": download_url,
                "file_name": file_name
            }],
            # 兼容旧格式
            success=True,
            message="文档生成成功",
            download_url=download_url,
            file_name=file_name
        )
        
    except HTTPException:
        # 重新抛出HTTP异常
        raise
    except Exception as e:
        logger.error(f"生成文档时发生未知错误: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"生成文档时发生错误: {str(e)}"
        )

@app.post(
    "/upload_file",
    responses={
        200: {"description": "文件上传成功"},
        400: {"model": ErrorResponse},
        401: {"model": ErrorResponse},
        500: {"model": ErrorResponse}
    }
)
async def upload_file(
    files: List[UploadFile] = File(None, description="上传的文件列表，最多3个"),
    attachments: List[UploadFile] = File(None, description="Dify上传的附件列表，最多3个"),
    token_valid: bool = Depends(verify_token)
) -> Dict[str, Any]:
    """
    上传文件接口 - 处理 Dify 上传的文件（支持 multipart/form-data）
    
    Args:
        files: 上传的文件列表
        attachments: Dify上传的附件列表
        token_valid: Token验证结果
        
    Returns:
        Dict[str, Any]: 上传结果
    """
    try:
        # 确定使用哪个字段的文件
        upload_files = files if files else attachments
        
        if not upload_files:
            raise HTTPException(
                status_code=400,
                detail="没有上传任何文件"
            )
        
        logger.info(f"开始处理文件上传，文件数量: {len(upload_files)}")
        
        # 验证附件数量
        if len(upload_files) > 3:
            raise HTTPException(
                status_code=400,
                detail="附件数量不能超过3个"
            )
        
        # 处理上传的文件
        processed_attachments = []
        for i, file in enumerate(upload_files):
            try:
                # 读取文件内容
                content = await file.read()
                
                # 获取文件类型
                file_type = attachment_processor._get_file_type(file.filename)
                
                # 提取标题
                if file_type == 'word':
                    # 对Word文件进行标题提取
                    extracted_title = attachment_processor._extract_title_from_word(content, file.filename)
                else:
                    # 对非Word文件使用文件名作为标题
                    extracted_title = file.filename.rsplit('.', 1)[0] if '.' in file.filename else file.filename
                
                # 构造附件信息
                attachment_info = {
                    "order": str(i + 1),
                    "name": file.filename.rsplit('.', 1)[0] if '.' in file.filename else file.filename,
                    "type": file_type,
                    "content": content,
                    "filename": file.filename,
                    "size": len(content),
                    "title": extracted_title,
                    "extracted_title": extracted_title  # 添加extracted_title字段
                }
                
                processed_attachments.append(attachment_info)
                logger.info(f"处理文件: {file.filename}, 大小: {len(content)} bytes, 提取标题: {extracted_title}")
                
            except Exception as e:
                logger.error(f"处理文件 {file.filename} 时发生错误: {str(e)}")
                raise HTTPException(
                    status_code=400,
                    detail=f"处理文件 {file.filename} 时发生错误: {str(e)}"
                )
        
        if not processed_attachments:
            raise HTTPException(
                status_code=400,
                detail="没有成功处理的附件"
            )
        
        # 生成会话ID用于关联附件
        import uuid
        session_id = str(uuid.uuid4())
        
        # 存储处理后的附件
        uploaded_attachments[session_id] = processed_attachments
        
        logger.info(f"文件上传处理完成，会话ID: {session_id}")
        
        return {
            "success": True,
            "message": f"成功处理 {len(processed_attachments)} 个附件",
            "session_id": session_id,
            "attachments_count": len(processed_attachments),
            "attachments": [
                {
                    "order": att["order"],
                    "name": att["name"],
                    "type": att["type"],
                    "title": att.get("title", ""),
                    "extracted_title": att.get("extracted_title", ""),
                    "size": att.get("size", 0)
                }
                for att in processed_attachments
            ]
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"文件上传处理时发生错误: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"文件上传处理失败: {str(e)}"
        )

@app.post(
    "/generate_document_with_attachments",
    response_model=DocumentGenerateResponse,
    responses={
        400: {"model": ErrorResponse},
        401: {"model": ErrorResponse},
        500: {"model": ErrorResponse}
    }
)
async def generate_document_with_attachments(
    request: DocumentGenerateWithAttachmentsRequest,
    session_id: str,
    token_valid: bool = Depends(verify_token)
) -> DocumentGenerateResponse:
    """
    生成带附件的公文文档
    
    Args:
        request: 文档生成请求
        session_id: 会话ID，用于获取上传的附件
        token_valid: Token验证结果
        
    Returns:
        DocumentGenerateResponse: 文档生成响应
    """
    try:
        logger.info(f"开始生成带附件的公文: {request.title}")
        
        # 验证请求参数
        if not request.content or not request.content.strip():
            raise HTTPException(
                status_code=400,
                detail="正文内容不能为空"
            )
        
        if not request.title.strip():
            raise HTTPException(
                status_code=400,
                detail="文档标题不能为空"
            )
        
        if not request.issuing_department.strip():
            raise HTTPException(
                status_code=400,
                detail="发文部门不能为空"
            )
        
        if not request.issue_date.strip():
            raise HTTPException(
                status_code=400,
                detail="发文日期不能为空"
            )
        
        # 获取上传的附件
        if session_id not in uploaded_attachments:
            raise HTTPException(
                status_code=400,
                detail="未找到对应的附件，请先上传附件"
            )
        
        attachments_data = uploaded_attachments[session_id]
        
        # 生成Word文档
        logger.info("开始生成Word文档")
        document_bytes = document_generator.generate_document(
            title=request.title,
            issuing_department=request.issuing_department,
            issue_date=request.issue_date,
            content=request.content,
            receiving_department=request.receiving_department,
            has_attachments=True,
            attachments=attachments_data
        )
        
        # 上传到OSS
        logger.info("开始上传文档到OSS")
        
        if oss_service is None:
            logger.error("OSS服务未初始化")
            raise HTTPException(
                status_code=500,
                detail="OSS服务未初始化，无法上传文档"
            )
        
        success, message, download_url = oss_service.upload_document(
            file_content=document_bytes,
            title=request.title,
            issue_date=request.issue_date
        )
        
        if not success:
            logger.error(f"上传文档失败: {message}")
            raise HTTPException(
                status_code=500,
                detail=f"上传文档失败: {message}"
            )
        
        # 生成文件名
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = request.title.replace('/', '_').replace('\\', '_')
        file_name = f"{timestamp}_{safe_title}.docx"
        
        logger.info(f"带附件文档生成成功: {file_name}")
        
        # 清理会话数据
        if session_id in uploaded_attachments:
            del uploaded_attachments[session_id]
        
        # 返回响应
        return DocumentGenerateResponse(
            body="文档生成成功",
            status_code=200,
            headers={
                "Content-Type": "application/json",
                "X-Generated-At": timestamp,
                "X-Service": "official-document-generator"
            },
            files=[{
                "download_url": download_url,
                "file_name": file_name
            }],
            # 兼容旧格式
            success=True,
            message="文档生成成功",
            download_url=download_url,
            file_name=file_name
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"生成带附件文档时发生错误: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"生成带附件文档时发生错误: {str(e)}"
        )

@app.post(
    "/generate_document_without_attachments",
    response_model=DocumentGenerateResponse,
    responses={
        400: {"model": ErrorResponse},
        401: {"model": ErrorResponse},
        500: {"model": ErrorResponse}
    }
)
async def generate_document_without_attachments(
    request: DocumentGenerateWithoutAttachmentsRequest,
    token_valid: bool = Depends(verify_token)
) -> DocumentGenerateResponse:
    """
    生成无附件的公文
    
    Args:
        request: 无附件公文生成请求
        token_valid: Token验证结果
        
    Returns:
        DocumentGenerateResponse: 生成结果
    """
    try:
        logger.info(f"开始生成无附件公文: {request.title}")
        
        # 构造请求数据
        doc_request = DocumentGenerateRequest(
            content=request.content,
            title=request.title,
            issuing_department=request.issuing_department,
            issue_date=request.issue_date,
            receiving_department=request.receiving_department,
            has_attachments=False,
            attachments=[]
        )
        
        # 生成文档
        document_bytes = document_generator.generate_document(
            title=doc_request.title,
            issuing_department=doc_request.issuing_department,
            issue_date=doc_request.issue_date,
            content=doc_request.content,
            receiving_department=doc_request.receiving_department,
            has_attachments=doc_request.has_attachments,
            attachments=doc_request.attachments
        )
        
        # 上传到OSS
        logger.info("开始上传文档到OSS")
        success, message, download_url = oss_service.upload_document(document_bytes, doc_request.title, doc_request.issue_date)
        
        if not success:
            raise HTTPException(
                status_code=500,
                detail=f"文档上传失败: {message}"
            )
        
        # 构造文件信息
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = doc_request.title.replace("/", "_").replace("\\", "_")
        filename = f"{timestamp}_{safe_title}.docx"
        
        file_info = {
            "filename": filename,
            "download_url": download_url,
            "size": len(document_bytes),
            "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        }
        
        logger.info(f"无附件公文生成成功: {filename}")
        
        return DocumentGenerateResponse(
            body="文档生成成功",
            status_code=200,
            headers={
                "Content-Type": "application/json",
                "X-Generated-At": datetime.now().strftime("%Y%m%d_%H%M%S"),
                "X-Service": "official-document-generator"
            },
            files=[
                FileInfo(
                    download_url=download_url,
                    file_name=filename
                )
            ],
            # 兼容字段
            success=True,
            message="文档生成成功",
            download_url=download_url,
            file_name=filename
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"生成无附件公文时发生错误: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"生成公文失败: {str(e)}"
        )

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    """全局异常处理器"""
    logger.error(f"全局异常: {str(exc)}")
    return JSONResponse(
        status_code=500,
        content={
            "success": False,
            "message": "服务器内部错误",
            "error_code": "INTERNAL_SERVER_ERROR"
        }
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "app.main:app",
        host=settings.APP_HOST,
        port=settings.APP_PORT,
        reload=settings.DEBUG
    ) 