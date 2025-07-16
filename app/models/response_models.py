"""
响应数据模型
"""
from pydantic import BaseModel, Field
from typing import Optional, Dict, Any, List

class FileInfo(BaseModel):
    """文件信息模型"""
    download_url: str = Field(..., description="文件下载链接")
    file_name: str = Field(..., description="文件名")

class DocumentGenerateResponse(BaseModel):
    """公文生成响应模型"""
    
    # 新格式字段
    body: str = Field(..., description="响应内容")
    status_code: int = Field(..., description="响应状态码")
    headers: Dict[str, Any] = Field(..., description="响应头列表JSON")
    files: List[FileInfo] = Field(..., description="文件列表，下载链接和文件名在文件列表里")
    
    # 兼容旧格式字段
    success: Optional[bool] = Field(None, description="是否成功(兼容字段)")
    message: Optional[str] = Field(None, description="响应信息(兼容字段)")
    download_url: Optional[str] = Field(None, description="文档下载链接(兼容字段)")
    file_name: Optional[str] = Field(None, description="生成的文件名(兼容字段)")
    
    class Config:
        """配置"""
        json_schema_extra = {
            "example": {
                "body": "文档生成成功",
                "status_code": 200,
                "headers": {
                    "Content-Type": "application/json",
                    "X-Generated-At": "2024-01-15T10:00:00Z"
                },
                "files": [
                    {
                        "download_url": "https://official_document.oss-cn-shanghai.aliyuncs.com/official_documents/20240115_关于加强公文写作规范的通知.docx",
                        "file_name": "20240115_关于加强公文写作规范的通知.docx"
                    }
                ]
            }
        }

class ErrorResponse(BaseModel):
    """错误响应模型"""
    
    success: bool = Field(False, description="是否成功")
    message: str = Field(..., description="错误信息")
    error_code: Optional[str] = Field(None, description="错误代码")
    
    class Config:
        """配置"""
        json_schema_extra = {
            "example": {
                "success": False,
                "message": "参数验证失败",
                "error_code": "VALIDATION_ERROR"
            }
        } 