"""
阿里云OSS服务
"""
import os
import oss2
import logging
from datetime import datetime
from typing import Optional, Tuple
from io import BytesIO

from app.config import settings

logger = logging.getLogger(__name__)

class OSSService:
    """阿里云OSS服务类"""
    
    def __init__(self):
        """初始化OSS客户端"""
        self.bucket = None
        self._initialize_oss()
    
    def _initialize_oss(self):
        """初始化OSS连接"""
        try:
            logger.info("开始初始化OSS服务...")
            logger.info(f"OSS配置 - Endpoint: {settings.OSS_ENDPOINT}")
            logger.info(f"OSS配置 - Bucket: {settings.OSS_BUCKET_NAME}")
            logger.info(f"OSS配置 - AccessKeyId: {settings.OSS_ACCESS_KEY_ID[:8]}***")
            
            # 创建OSS认证对象
            auth = oss2.Auth(settings.OSS_ACCESS_KEY_ID, settings.OSS_ACCESS_KEY_SECRET)
            
            # 创建Bucket对象
            self.bucket = oss2.Bucket(auth, settings.OSS_ENDPOINT, settings.OSS_BUCKET_NAME)
            
            logger.info("OSS服务初始化成功")
            
        except Exception as e:
            logger.error(f"OSS服务初始化失败: {str(e)}")
            logger.error(f"请检查OSS配置信息是否正确")
            # 不抛出异常，允许服务启动，但OSS功能不可用
            self.bucket = None
    
    def upload_document(self, file_content: bytes, title: str, issue_date: str) -> Tuple[bool, str, Optional[str]]:
        """
        上传文档到OSS
        
        Args:
            file_content: 文件内容（字节）
            title: 文档标题
            issue_date: 发文日期
            
        Returns:
            Tuple[bool, str, Optional[str]]: (是否成功, 消息, 下载链接)
        """
        if not self.bucket:
            logger.error("OSS未正确初始化，无法上传文档")
            return False, "OSS服务未正确初始化", None
            
        try:
            # 生成文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_title = self._sanitize_filename(title)
            filename = f"{timestamp}_{safe_title}.docx"
            
            # OSS对象键（完整路径）
            object_key = f"{settings.OSS_DOCUMENT_PREFIX}{filename}"
            
            logger.info(f"开始上传文档到OSS: {object_key}")
            
            # 设置上传参数，添加正确的Content-Type和Content-Disposition
            # 使用RFC 5987格式支持中文文件名
            import urllib.parse
            encoded_filename = urllib.parse.quote(filename, safe='')
            headers = {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition': f'attachment; filename*=UTF-8\'\'{encoded_filename}'
            }
            
            # 上传文件
            result = self.bucket.put_object(object_key, file_content, headers=headers)
            
            if result.status == 200:
                # 生成带签名的下载链接（有效期7天）
                # 确保对象键没有重复编码
                import urllib.parse
                unquoted_key = urllib.parse.unquote(object_key)
                download_url = self.bucket.sign_url('GET', unquoted_key, 7*24*3600)
                
                # 修复URL编码问题
                if '%2F' in download_url:
                    download_url = download_url.replace('%2F', '/')
                
                logger.info(f"文档上传成功: {filename}")
                logger.info(f"对象键: {object_key}")
                logger.info(f"解码后对象键: {unquoted_key}")
                logger.info(f"下载链接: {download_url}")
                return True, "文档上传成功", download_url
            else:
                logger.error(f"文档上传失败，状态码: {result.status}")
                return False, f"文档上传失败，状态码: {result.status}", None
                
        except oss2.exceptions.OssError as e:
            logger.error(f"OSS错误: {e.code} - {e.message}")
            return False, f"OSS错误: {e.code} - {e.message}", None
        except Exception as e:
            logger.error(f"上传文档时发生错误: {str(e)}")
            return False, f"上传文档时发生错误: {str(e)}", None
    
    def _sanitize_filename(self, filename: str) -> str:
        """
        清理文件名，移除不安全字符
        
        Args:
            filename: 原始文件名
            
        Returns:
            str: 清理后的文件名
        """
        # 移除或替换不安全字符
        unsafe_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '\n', '\r', '\t']
        safe_filename = filename
        
        for char in unsafe_chars:
            safe_filename = safe_filename.replace(char, '_')
        
        # 限制长度
        if len(safe_filename) > 100:
            safe_filename = safe_filename[:100]
        
        return safe_filename
    
    def check_bucket_exists(self) -> bool:
        """
        检查Bucket是否存在
        
        Returns:
            bool: Bucket是否存在
        """
        if not self.bucket:
            logger.warning("OSS未正确初始化，无法检查Bucket")
            return False
            
        try:
            # 使用正确的API检查bucket是否存在
            self.bucket.get_bucket_info()
            logger.info(f"Bucket {settings.OSS_BUCKET_NAME} 连接正常")
            return True
        except oss2.exceptions.NoSuchBucket:
            logger.warning(f"Bucket {settings.OSS_BUCKET_NAME} 不存在")
            return False
        except oss2.exceptions.AccessDenied:
            logger.warning(f"Bucket {settings.OSS_BUCKET_NAME} 访问被拒绝，请检查权限")
            return False
        except oss2.exceptions.OssError as e:
            logger.error(f"检查Bucket时发生OSS错误: {e.code} - {e.message}")
            return False
        except Exception as e:
            logger.error(f"检查Bucket时发生错误: {str(e)}")
            return False

# 全局OSS服务实例
try:
    oss_service = OSSService()
except Exception as e:
    logger.error(f"创建OSS服务实例失败: {str(e)}")
    # 创建一个空的服务实例
    oss_service = None 