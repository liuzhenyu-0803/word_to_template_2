#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
模型管理器 - 提供统一的LLM调用接口
支持本地模型和远程API两种方式
"""

import os
import time
from typing import List, Dict, Any, Optional, Union

# 本地模型相关
try:
    from llama_cpp import Llama
    LLAMA_CPP_AVAILABLE = True
except ImportError:
    LLAMA_CPP_AVAILABLE = False

# OpenAI SDK相关
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False


class ModelManager:
    """
    模型管理器类 - 单例模式
    提供统一的接口调用LLM，支持本地模型和远程API
    """
    
    def __init__(self):
        # 初始化状态
        self.model_type = None
        self.local_model = None
        self.remote_client = None
        self.remote_model = None
        self.api_key = None
        self.base_url = None
        
        # 默认使用本地模型
        self.use_local_model = True
    
    def init_local_model(self, model_name: str = "Qwen3-0.6B-Q8_0.gguf") -> bool:
        """
        初始化本地模型
        
        Args:
            model_name: 模型文件名，默认为Qwen3-0.6B-Q8_0.gguf
            
        Returns:
            bool: 是否成功初始化
        """
        if not LLAMA_CPP_AVAILABLE:
            print("错误: llama_cpp 库未安装，无法使用本地模型")
            return False
            
        try:
            # 获取models目录路径
            current_dir = os.path.dirname(os.path.abspath(__file__))
            model_name = "gemma-3-4b-it-Q4_K_M.gguf"
            model_path = os.path.join(current_dir, model_name)
            
            # 检查模型文件是否存在
            if not os.path.exists(model_path):
                print(f"错误: 模型文件不存在: {model_path}")
                return False
                
            # 加载模型，使用更高效的CPU优化参数
            import multiprocessing
            logical_cores = multiprocessing.cpu_count()  # 获取CPU逻辑核心数
            physical_cores = max(1, logical_cores // 3 * 2)  # 估算物理核心数
            
            self.local_model = Llama(
                n_gpu_layers=34,
                model_path=model_path, 
                n_ctx=32768,               # 缩小上下文窗口以提高性能
                # use_mmap=True,           # 启用内存映射
                # flash_attn=True,         # 启用FlashAttention
                # n_threads = 16,  
                # n_threads_batch = 16,  # 批处理时使用的线程数
                # n_batch = 1024,           # 批处理大小
                # n_ubatch = 1024,           # 批处理大小
                # use_mlock = True,        # 启用内存锁定
            )
            self.model_type = "local"
            self.use_local_model = True
            print(f"成功加载本地模型: {model_name}")
            return True
            
        except Exception as e:
            print(f"初始化本地模型失败: {str(e)}")
            return False
    
    def init_remote_model(self, 
                         api_key: str = "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",
                         base_url: str = "https://openrouter.ai/api/v1",
                         model: str = "google/gemma-3-4b-it") -> bool:
        """
        初始化远程API模型
        
        Args:
            api_key: API密钥
            base_url: API基础URL
            model: 模型名称
            
        Returns:
            bool: 是否成功初始化
        """
        if not OPENAI_AVAILABLE:
            print("错误: openai 库未安装，无法使用远程API")
            return False

        try:
            # 初始化客户端

            model = "google/gemma-3-4b-it"

            # lm studio
            # base_url = "http://127.0.0.1:1234/v1"
            # model = "gemma-3-4b-it-qat"
            
            self.remote_client = OpenAI(api_key=api_key, base_url=base_url)
            self.base_url = base_url
            self.api_key = api_key
            self.model_type = "remote"
            self.use_local_model = False
            self.remote_model = model
            print(f"成功初始化远程模型: {model}")
            return True
            
        except Exception as e:
            print(f"初始化远程模型失败: {str(e)}")
            return False
    
    def create_completion(self, 
                         messages: List[Dict[str, str]], 
                         temperature: float = 0, 
                         max_retries: int = 3) -> Optional[str]:
        """
        创建聊天完成
        
        Args:
            messages: 消息列表，格式为[{"role": "user", "content": "内容"}]
            temperature: 采样温度
            max_retries: 最大重试次数
            
        Returns:
            str: 模型返回的内容，失败时返回None
        """
        # 检查模型是否已初始化
        if self.use_local_model and self.local_model is None:
            print("错误: 本地模型未初始化")
            return None
            
        if not self.use_local_model and self.remote_client is None:
            print("错误: 远程模型未初始化")
            return None
            
        # 根据模型类型调用不同的API
        for attempt in range(max_retries):
            try:
                if self.use_local_model:
                    return self._call_local_model(messages, temperature)
                else:
                    return self._call_remote_model(messages, temperature)
                    
            except Exception as e:
                print(f"尝试 {attempt+1}/{max_retries}: 调用模型失败 - {str(e)}")
                if attempt < max_retries - 1:
                    print("将在2秒后重试...")
                    time.sleep(2)
                else:
                    print(f"已达到最大重试次数 ({max_retries})，放弃")
                    return None
    
    def _call_local_model(self, messages: List[Dict[str, str]], temperature: float) -> str:
        """调用本地模型"""
        response = self.local_model.create_chat_completion(
            messages=messages,
            temperature=temperature
        )
        return response["choices"][0]["message"]["content"]
    
    def _call_remote_model(self, messages: List[Dict[str, str]], temperature: float) -> str:
        """调用远程API模型"""
        response = self.remote_client.chat.completions.create(
            model=self.remote_model,
            messages=messages,
            temperature=temperature
        )
        return response.choices[0].message.content

    # 注意：删除了get_model_info方法，因为项目中没有使用到


# 全局单例实例 - 直接在模块级别创建，其他模块导入时自动使用同一个实例
llm_manager = ModelManager()

# 示例用法
if __name__ == "__main__":
    # # 初始化本地模型
    success = llm_manager.init_local_model()
    
    # if success:
    #     # 测试本地模型
    #     response = llm_manager.create_completion([
    #         {"role": "user", "content": "你好，介绍一下自己"}
    #     ])
    #     print("本地模型响应:")
    #     print(response)
    
    # 初始化远程模型
    # success = llm_manager.init_remote_model()
    
    if success:
        # 测试远程模型
        response = llm_manager.create_completion([
            {"role": "user", "content": "你好，介绍一下自己"}
        ])
        print("\n远程模型响应:")
        print(response)
