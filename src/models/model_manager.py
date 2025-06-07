#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
模型管理器 - 提供统一的LLM调用接口
支持本地模型和远程API两种方式
"""

import os
from typing import List, Dict, Any, Optional, Union

from llama_cpp import Llama
from openai import OpenAI


class ModelManager:
    """
    模型管理器类 - 单例模式
    提供统一的接口调用LLM，支持本地模型和远程API
    """
    
    def __init__(self):
        # 初始化状态
        self.local_model = None
        self.remote_client = None
        self.remote_model = None
        self.api_key = None
        self.base_url = None
        
        # 默认使用本地模型
        self.use_local_model = True
    
    def init_local_model(self, model_name: str = "gemma-3-4b-it-Q4_K_M.gguf") -> bool:
        """
        初始化本地模型
        
        Args:
            model_name: 模型文件名，默认为gemma-3-4b-it-Q4_K_M.gguf

        Returns:
            bool: 是否成功初始化
        """

        try:
            # 获取models目录路径
            current_dir = os.path.dirname(os.path.abspath(__file__))
            model_name = "gemma-3-4b-it-Q4_K_M.gguf"
            model_path = os.path.join(current_dir, model_name)
            
            # 检查模型文件是否存在
            if not os.path.exists(model_path):
                print(f"错误: 模型文件不存在: {model_path}")
                return False
                
            self.local_model = Llama(
                model_path=model_path, 
                n_ctx=10000,
                # offload_kqv=False,
                # use_mmap=False,
                # flash_attn=True,
            )
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

        try:
            # 初始化客户端            
            self.remote_client = OpenAI(api_key=api_key, base_url=base_url)
            self.base_url = base_url
            self.api_key = api_key
            self.use_local_model = False
            self.remote_model = model
            print(f"成功初始化远程模型: {model}")
            return True
            
        except Exception as e:
            print(f"初始化远程模型失败: {str(e)}")
            return False
    
    def create_completion(self,
                         messages: List[Dict[str, str]], 
                         temperature: float = 0) -> Optional[str]:
        """
        创建聊天完成
        
        Args:
            messages: 消息列表，格式为[{"role": "user", "content": "内容"}]
            temperature: 采样温度
            
        Returns:
            str: 模型返回的内容，失败时返回None
        """

        # 根据模型类型调用不同的API
        try:
            if self.use_local_model:
                return self._call_local_model(messages, temperature)
            else:
                return self._call_remote_model(messages, temperature)
                
        except Exception as e:
            print(f"调用模型失败: {str(e)}")
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


# 全局单例实例 - 直接在模块级别创建，其他模块导入时自动使用同一个实例
llm_manager = ModelManager()

# 示例用法
if __name__ == "__main__":
    # 初始化本地模型
    success = llm_manager.init_local_model()
    
    if success:
        # 测试本地模型
        response = llm_manager.create_completion([
            {"role": "user", "content": "你好，介绍一下自己"}
        ])
        print("本地模型响应:")
        print(response)
    
    # 初始化远程模型
    # success = llm_manager.init_remote_model()
    
    # if success:
    #     # 测试远程模型
    #     response = llm_manager.create_completion([
    #         {"role": "user", "content": "你好，介绍一下自己"}
    #     ])
    #     print("\n远程模型响应:")
    #     print(response)
