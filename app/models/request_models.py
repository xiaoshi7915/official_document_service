"""
请求模型定义
"""
from pydantic import BaseModel, Field
from typing import List, Optional, Any, Dict
from datetime import datetime

class AttachmentModel(BaseModel):
    """附件模型"""
    order: str = Field(..., description="附件顺序号，如：1、2、3")
    type: str = Field(..., description="附件类型：csv、table、text等")
    name: str = Field(..., description="附件名称")
    markdown_content: str = Field(..., description="附件的markdown格式内容")
    
class DifyAttachmentModel(BaseModel):
    """Dify 附件模型"""
    dify_model_identity: str = Field(..., description="Dify 模型标识")
    id: Optional[str] = Field(None, description="附件ID")
    tenant_id: str = Field(..., description="租户ID")
    type: str = Field(..., description="附件类型")
    transfer_method: str = Field(..., description="传输方法")
    remote_url: str = Field(..., description="远程URL")
    related_id: str = Field(..., description="关联ID")
    filename: str = Field(..., description="文件名")
    extension: str = Field(..., description="文件扩展名")
    mime_type: str = Field(..., description="MIME类型")
    size: int = Field(..., description="文件大小")
    url: str = Field(..., description="文件URL")

class DocumentGenerateRequest(BaseModel):
    """公文生成请求模型"""
    
    # 支持两种输入格式的字段
    content: Optional[str] = Field(None, description="markdown格式的正文内容(旧格式)")
    markdown_content: Optional[str] = Field(None, description="markdown格式的正文内容(新格式)")
    
    title: str = Field(..., description="文档标题")
    issuing_department: str = Field(..., description="发文部门")
    issue_date: str = Field(..., description="发文日期，格式：YYYY年MM月DD日")
    receiving_department: Optional[str] = Field(None, description="收文部门（人）")
    has_attachments: bool = Field(default=False, description="是否有附件")
    
    # 支持多种附件格式：对象数组(旧)、字符串数组(新)、空数组
    attachments: Optional[Any] = Field(
        default=None, 
        description="附件列表，支持字符串数组或对象数组，最多3个附件"
    )
    
    class Config:
        """配置"""
        json_schema_extra = {
            "example": {
                "content": "这是一个示例公文内容...",
                "title": "关于XX工作的通知",
                "issuing_department": "XX部门",
                "issue_date": "2024年1月1日",
                "receiving_department": "XX单位",
                "has_attachments": False,
                "attachments": []
            }
        }

class DocumentGenerateWithAttachmentsRequest(BaseModel):
    """带附件的公文生成请求模型"""
    
    content: str = Field(..., description="markdown格式的正文内容")
    title: str = Field(..., description="文档标题")
    issuing_department: str = Field(..., description="发文部门")
    issue_date: str = Field(..., description="发文日期，格式：YYYY年MM月DD日")
    receiving_department: Optional[str] = Field(None, description="收文部门（人）")
    has_attachments: bool = Field(default=True, description="是否有附件，固定为true")
    
    class Config:
        """配置"""
        json_schema_extra = {
            "example": {
                "content": "这是一个示例公文内容...",
                "title": "关于XX工作的通知",
                "issuing_department": "XX部门",
                "issue_date": "2024年1月1日",
                "receiving_department": "XX单位",
                "has_attachments": True
            }
        }

class DocumentGenerateWithoutAttachmentsRequest(BaseModel):
    """无附件的公文生成请求模型"""
    
    content: str = Field(..., description="markdown格式的正文内容")
    title: str = Field(..., description="文档标题")
    issuing_department: str = Field(..., description="发文部门")
    issue_date: str = Field(..., description="发文日期，格式：YYYY年MM月DD日")
    receiving_department: Optional[str] = Field(None, description="收文部门（人）")
    has_attachments: bool = Field(default=False, description="是否有附件，固定为false")
    
    class Config:
        """配置"""
        json_schema_extra = {
            "example": {
                "content": "这是一个示例公文内容...",
                "title": "关于XX工作的通知",
                "issuing_department": "XX部门",
                "issue_date": "2024年1月1日",
                "receiving_department": "XX单位",
                "has_attachments": False
            }
        }

class FileUploadRequest(BaseModel):
    """文件上传请求模型 - 只用于附件上传"""
    
    attachments: List[DifyAttachmentModel] = Field(..., description="Dify 附件列表，最多3个")
    
    class Config:
        """配置"""
        json_schema_extra = {
            "example": {
                "attachments": [
                    {
                        "dify_model_identity": "__dify__file__",
                        "id": None,
                        "tenant_id": "tenant123",
                        "type": "document",
                        "transfer_method": "local_file",
                        "remote_url": "",
                        "related_id": "file123",
                        "filename": "example.docx",
                        "extension": ".docx",
                        "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "size": 12345,
                        "url": "https://example.com/files/file123"
                    }
                ]
            }
        } 