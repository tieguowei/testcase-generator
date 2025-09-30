#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试用例转思维导图工具 - 简化版
将Tab缩进的测试用例转换为最简化的XMind格式思维导图
"""

import os
from pathlib import Path
import zipfile


class SimpleTestCaseToXMindConverter:
    """简化版测试用例转XMind转换器"""

    def __init__(self):
        self.node_id_counter = 0

    def convert(self, input_file: str, output_file: str = None):
        """
        转换测试用例为XMind格式

        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径（可选）
        """
        try:
            # 读取测试用例文件
            with open(input_file, 'r', encoding='utf-8') as f:
                content = f.read()

            # 解析内容
            root_node = self._parse_content(content)

            # 生成输出文件名
            if not output_file:
                input_path = Path(input_file)
                output_file = input_path.parent / f"{input_path.stem}.xmind"

            # 创建XMind文件
            self._create_simple_xmind_file(root_node, output_file)

            print(f"转换完成！输出文件: {output_file}")
            return str(output_file)

        except Exception as e:
            print(f"转换过程中出错: {str(e)}")
            raise

    def _parse_content(self, content: str):
        """解析测试用例内容并构建层级结构"""
        lines = content.strip().split('\n')
        
        # 创建根节点
        root_node = {
            'id': self._get_next_id(),
            'title': '测试用例',
            'children': []
        }
        
        # 用栈来跟踪当前路径
        node_stack = [root_node]
        
        for line in lines:
            if not line.strip():
                continue
                
            # 计算缩进级别
            indent_level = self._get_indent_level(line)
            text = line.strip()
            
            # 清理文本
            clean_text = self._clean_text(text)
            
            # 调整栈的深度到正确的层级
            while len(node_stack) > indent_level + 1:
                node_stack.pop()
            
            # 创建新节点
            new_node = {
                'id': self._get_next_id(),
                'title': clean_text,
                'children': []
            }
            
            # 添加到当前父节点
            node_stack[-1]['children'].append(new_node)
            
            # 将新节点加入栈
            node_stack.append(new_node)
        
        return root_node

    def _create_simple_xmind_file(self, root_node, output_file):
        """创建最简化的XMind文件"""
        
        # 逻辑图格式的content.xml
        content_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<xmap-content xmlns="urn:xmind:xmap:xmlns:content:2.0" version="2.0">
  <sheet id="sheet1">
    <topic id="{root_node['id']}" structure-class="org.xmind.ui.logic.right">
      <title>{root_node['title']}</title>
{self._build_simple_children_xml(root_node['children'], 6)}
    </topic>
  </sheet>
</xmap-content>'''

        # 最简化的meta.xml
        meta_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<meta xmlns="urn:xmind:xmap:xmlns:meta:2.0" version="2.0">
</meta>'''

        # 最简化的manifest.xml
        manifest_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest xmlns="urn:xmind:xmap:xmlns:manifest:1.0">
  <file-entry full-path="content.xml" media-type="text/xml"/>
  <file-entry full-path="meta.xml" media-type="text/xml"/>
</manifest>'''

        # 创建XMind文件
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as xmind_zip:
            xmind_zip.writestr('content.xml', content_xml)
            xmind_zip.writestr('meta.xml', meta_xml)
            xmind_zip.writestr('META-INF/manifest.xml', manifest_xml)

    def _build_simple_children_xml(self, children, indent_level):
        """递归构建子节点XML - 简化版"""
        if not children:
            return ""
        
        indent = " " * indent_level
        child_indent = " " * (indent_level + 2)
        topic_indent = " " * (indent_level + 4)
        
        xml_parts = [f"{indent}<children>"]
        xml_parts.append(f"{child_indent}<topics type=\"attached\">")
        
        for child in children:
            xml_parts.append(f"{topic_indent}<topic id=\"{child['id']}\">")
            xml_parts.append(f"{topic_indent}  <title>{child['title']}</title>")
            
            if child['children']:
                xml_parts.append(self._build_simple_children_xml(child['children'], indent_level + 6))
            
            xml_parts.append(f"{topic_indent}</topic>")
        
        xml_parts.append(f"{child_indent}</topics>")
        xml_parts.append(f"{indent}</children>")
        
        return "\n".join(xml_parts)

    def _get_indent_level(self, line: str) -> int:
        """获取行的缩进级别"""
        tab_count = 0
        for char in line:
            if char == '\t':
                tab_count += 1
            else:
                break
        return tab_count

    def _clean_text(self, text: str) -> str:
        """清理文本，移除前缀并转义特殊字符"""
        # if text.startswith('测试点：'):
        #     text = text[4:]
        # elif text.startswith('用例步骤：'):
        #     text = text[5:]
        # elif text.startswith('预期结果：'):
        #     text = text[5:]

        # 转义XML特殊字符
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        text = text.replace("'", '&apos;')

        return text

    def _get_next_id(self) -> str:
        """获取下一个节点ID"""
        self.node_id_counter += 1
        return f"topic{self.node_id_counter}"


def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description='测试用例转思维导图工具 - 简化版')
    parser.add_argument('input', help='输入的测试用例文件路径')
    parser.add_argument('-o', '--output', help='输出文件路径')

    args = parser.parse_args()

    # 检查输入文件
    if not os.path.exists(args.input):
        print(f"错误: 文件 '{args.input}' 不存在")
        return 1

    # 创建转换器
    converter = SimpleTestCaseToXMindConverter()

    try:
        # 执行转换
        output_path = converter.convert(args.input, args.output)
        print(f"转换成功! 输出文件: {output_path}")
        return 0
    except Exception as e:
        print(f"转换失败: {str(e)}")
        return 1


if __name__ == '__main__':
    exit(main())
