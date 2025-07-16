#!/usr/bin/env python3
"""
公文写作API服务启动脚本
"""
import uvicorn
from app.config import settings

if __name__ == "__main__":
    uvicorn.run(
        "app.main:app",
        host=settings.APP_HOST,
        port=settings.APP_PORT,
        reload=settings.DEBUG,
        access_log=True,
        log_level="info"
    ) 