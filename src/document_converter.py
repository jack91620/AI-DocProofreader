#!/usr/bin/env python3
"""
文档格式转换器
支持.doc到.docx的转换，以及其他格式的处理
"""

import os
import sys
import subprocess
import shutil
import tempfile
from pathlib import Path
from typing import Optional, Tuple, Union
import platform


class DocumentConverter:
    """文档格式转换器"""
    
    def __init__(self):
        self.temp_dir = None
        self.conversion_method = self._detect_conversion_method()
    
    def _detect_conversion_method(self) -> str:
        """检测可用的转换方法"""
        # 1. 检查LibreOffice
        if self._check_libreoffice():
            return "libreoffice"
        
        # 2. 检查pandoc
        if self._check_pandoc():
            return "pandoc"
        
        # 3. Windows系统检查COM
        if platform.system() == "Windows":
            try:
                import win32com.client
                return "com"
            except ImportError:
                pass
        
        # 4. 检查在线转换服务
        if self._check_online_service():
            return "online"
        
        return "none"
    
    def _check_libreoffice(self) -> bool:
        """检查LibreOffice是否可用"""
        try:
            # 尝试不同的LibreOffice命令名称
            commands = ['libreoffice', 'soffice', '/usr/bin/libreoffice', 
                       '/Applications/LibreOffice.app/Contents/MacOS/soffice']
            
            for cmd in commands:
                try:
                    result = subprocess.run([cmd, '--version'], 
                                          capture_output=True, text=True, timeout=10)
                    if result.returncode == 0:
                        self.libreoffice_cmd = cmd
                        return True
                except (subprocess.TimeoutExpired, FileNotFoundError):
                    continue
            return False
        except Exception:
            return False
    
    def _check_pandoc(self) -> bool:
        """检查pandoc是否可用"""
        try:
            result = subprocess.run(['pandoc', '--version'], 
                                  capture_output=True, text=True, timeout=10)
            return result.returncode == 0
        except (subprocess.TimeoutExpired, FileNotFoundError):
            return False
    
    def _check_online_service(self) -> bool:
        """检查在线转换服务是否可用"""
        try:
            import requests
            # 简单的网络连接测试
            response = requests.get('https://httpbin.org/status/200', timeout=5)
            return response.status_code == 200
        except:
            return False
    
    def can_convert_doc(self) -> Tuple[bool, str]:
        """检查是否可以转换.doc文件"""
        if self.conversion_method == "none":
            return False, "未找到可用的文档转换工具"
        
        method_descriptions = {
            "libreoffice": "使用LibreOffice进行转换",
            "pandoc": "使用Pandoc进行转换", 
            "com": "使用Windows COM接口转换",
            "online": "使用在线转换服务"
        }
        
        return True, method_descriptions.get(self.conversion_method, "未知转换方法")
    
    def convert_doc_to_docx(self, doc_path: str, output_path: str = None) -> Optional[str]:
        """
        将.doc文件转换为.docx格式
        
        Args:
            doc_path: .doc文件路径
            output_path: 输出.docx文件路径，如果为None则自动生成
            
        Returns:
            转换成功返回.docx文件路径，失败返回None
        """
        if not os.path.exists(doc_path):
            print(f"❌ 源文件不存在: {doc_path}")
            return None
            
        if not doc_path.lower().endswith('.doc'):
            print(f"❌ 不是.doc文件: {doc_path}")
            return None
        
        if output_path is None:
            output_path = doc_path.rsplit('.', 1)[0] + '.docx'
        
        print(f"🔄 开始转换 {Path(doc_path).name} -> {Path(output_path).name}")
        print(f"📋 使用转换方法: {self.conversion_method}")
        
        try:
            if self.conversion_method == "libreoffice":
                return self._convert_with_libreoffice(doc_path, output_path)
            elif self.conversion_method == "pandoc":
                return self._convert_with_pandoc(doc_path, output_path)
            elif self.conversion_method == "com":
                return self._convert_with_com(doc_path, output_path)
            elif self.conversion_method == "online":
                return self._convert_with_online_service(doc_path, output_path)
            else:
                print(f"❌ 不支持的转换方法: {self.conversion_method}")
                return None
                
        except Exception as e:
            print(f"❌ 转换失败: {e}")
            return None
    
    def _convert_with_libreoffice(self, doc_path: str, output_path: str) -> Optional[str]:
        """使用LibreOffice转换"""
        try:
            output_dir = os.path.dirname(os.path.abspath(output_path))
            
            # LibreOffice转换命令
            cmd = [
                self.libreoffice_cmd,
                '--headless',
                '--convert-to', 'docx',
                '--outdir', output_dir,
                doc_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            if result.returncode == 0:
                # LibreOffice会自动生成文件名，需要重命名
                auto_generated = os.path.join(output_dir, 
                    os.path.splitext(os.path.basename(doc_path))[0] + '.docx')
                
                if os.path.exists(auto_generated) and auto_generated != output_path:
                    shutil.move(auto_generated, output_path)
                
                if os.path.exists(output_path):
                    print(f"✅ LibreOffice转换成功: {output_path}")
                    return output_path
            
            print(f"❌ LibreOffice转换失败: {result.stderr}")
            return None
            
        except subprocess.TimeoutExpired:
            print("❌ LibreOffice转换超时")
            return None
        except Exception as e:
            print(f"❌ LibreOffice转换异常: {e}")
            return None
    
    def _convert_with_pandoc(self, doc_path: str, output_path: str) -> Optional[str]:
        """使用Pandoc转换"""
        try:
            cmd = ['pandoc', doc_path, '-o', output_path]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            if result.returncode == 0 and os.path.exists(output_path):
                print(f"✅ Pandoc转换成功: {output_path}")
                return output_path
            
            print(f"❌ Pandoc转换失败: {result.stderr}")
            return None
            
        except subprocess.TimeoutExpired:
            print("❌ Pandoc转换超时")
            return None
        except Exception as e:
            print(f"❌ Pandoc转换异常: {e}")
            return None
    
    def _convert_with_com(self, doc_path: str, output_path: str) -> Optional[str]:
        """使用Windows COM接口转换"""
        try:
            import win32com.client
            
            # 创建Word应用程序对象
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                # 打开.doc文件
                doc = word.Documents.Open(os.path.abspath(doc_path))
                
                # 另存为.docx格式 (格式代码16为docx)
                doc.SaveAs2(os.path.abspath(output_path), FileFormat=16)
                doc.Close()
                
                if os.path.exists(output_path):
                    print(f"✅ COM转换成功: {output_path}")
                    return output_path
                    
            finally:
                word.Quit()
            
            return None
            
        except Exception as e:
            print(f"❌ COM转换异常: {e}")
            return None
    
    def _convert_with_online_service(self, doc_path: str, output_path: str) -> Optional[str]:
        """使用在线转换服务（示例实现）"""
        print("⚠️ 在线转换服务需要上传文件，请确保文档不包含敏感信息")
        print("❌ 在线转换服务暂未实现")
        return None
    
    def detect_document_format(self, file_path: str) -> str:
        """检测文档格式"""
        if not os.path.exists(file_path):
            return "unknown"
        
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.docx':
            return 'docx'
        elif ext == '.doc':
            return 'doc'
        elif ext in ['.txt', '.text']:
            return 'txt'
        elif ext in ['.rtf']:
            return 'rtf'
        elif ext in ['.odt']:
            return 'odt'
        else:
            return 'unknown'
    
    def ensure_docx_format(self, file_path: str, temp_dir: str = None) -> Tuple[str, bool]:
        """
        确保文件为.docx格式，如果不是则转换
        
        Returns:
            (docx_file_path, is_converted)
        """
        format_type = self.detect_document_format(file_path)
        
        if format_type == 'docx':
            return file_path, False
        
        if format_type == 'doc':
            if temp_dir is None:
                temp_dir = tempfile.mkdtemp()
            
            temp_docx = os.path.join(temp_dir, 
                os.path.splitext(os.path.basename(file_path))[0] + '.docx')
            
            converted_path = self.convert_doc_to_docx(file_path, temp_docx)
            if converted_path:
                return converted_path, True
            else:
                raise RuntimeError(f"无法转换.doc文件: {file_path}")
        
        raise ValueError(f"不支持的文档格式: {format_type}")
    
    def cleanup_temp_files(self, file_path: str, was_converted: bool):
        """清理临时文件"""
        if was_converted and os.path.exists(file_path):
            try:
                os.remove(file_path)
                # 如果是临时目录中的文件，也清理目录
                temp_dir = os.path.dirname(file_path)
                if temp_dir.startswith(tempfile.gettempdir()):
                    shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception as e:
                print(f"⚠️ 清理临时文件失败: {e}")


def install_conversion_tools():
    """安装文档转换工具的指导"""
    print("📋 文档转换工具安装指导:")
    print()
    
    system = platform.system()
    
    if system == "Linux":
        print("🐧 Linux系统:")
        print("  sudo apt-get update")
        print("  sudo apt-get install libreoffice")
        print("  # 或者")
        print("  sudo apt-get install pandoc")
    
    elif system == "Darwin":  # macOS
        print("🍎 macOS系统:")
        print("  # 使用Homebrew:")
        print("  brew install --cask libreoffice")
        print("  # 或者")
        print("  brew install pandoc")
    
    elif system == "Windows":
        print("🪟 Windows系统:")
        print("  # 方法1: 下载LibreOffice")
        print("  https://www.libreoffice.org/download/download/")
        print("  # 方法2: 安装pywin32使用COM接口")
        print("  pip install pywin32")
        print("  # 方法3: 使用Chocolatey安装pandoc")
        print("  choco install pandoc")
    
    print()
    print("🔧 Python依赖:")
    print("  pip install python-docx")
    print("  # Windows用户额外安装:")
    print("  pip install pywin32")


if __name__ == "__main__":
    converter = DocumentConverter()
    can_convert, method = converter.can_convert_doc()
    
    print(f"📋 转换能力检测:")
    print(f"   可以转换.doc文件: {'是' if can_convert else '否'}")
    print(f"   转换方法: {method}")
    
    if not can_convert:
        print()
        install_conversion_tools()
    
    # 测试转换（如果有测试文件）
    test_doc = "test.doc"
    if os.path.exists(test_doc):
        result = converter.convert_doc_to_docx(test_doc)
        if result:
            print(f"✅ 测试转换成功: {result}")
        else:
            print("❌ 测试转换失败")
    else:
        print(f"ℹ️ 没有找到测试文件 {test_doc}") 