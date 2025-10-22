import json
import re
import os

def parse_translated_txt(txt_file_path):
    """
    解析翻译后的txt文件，提取每页的文本块内容
    
    Args:
        txt_file_path (str): 翻译后的txt文件路径
    
    Returns:
        dict: {page_number: [text_block1, text_block2, ...]}
    """
    try:
        with open(txt_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 按页面分割
        pages = re.split(r'=== 第(\d+)ページ ===', content)
        
        parsed_data = {}
        
        # 处理每一页（跳过第一个空元素）
        for i in range(1, len(pages), 2):
            if i + 1 < len(pages):
                page_number = int(pages[i])
                page_content = pages[i + 1].strip()
                
                # 提取文本块
                text_blocks = re.findall(r'テキストブロック \d+:\s*([\s\S]*?)(?=テキストブロック \d+:|$)', page_content)
                
                # 清理文本块内容
                cleaned_blocks = []
                for block in text_blocks:
                    cleaned_block = block.strip()
                    if cleaned_block:
                        cleaned_blocks.append(cleaned_block)
                
                parsed_data[page_number] = cleaned_blocks
        
        return parsed_data
    
    except Exception as e:
        print(f"❌ 解析翻译文件时出错: {e}")
        return {}

def fill_japanese_content(original_json_path, translated_txt_path, output_json_path):
    """
    将翻译后的文本内容填充到原始JSON结构中
    
    Args:
        original_json_path (str): 原始ppt_content.json文件路径
        translated_txt_path (str): 翻译后的txt文件路径
        output_json_path (str): 输出的日语JSON文件路径
    """
    try:
        # 读取原始JSON文件
        print("📖 读取原始JSON文件...")
        with open(original_json_path, 'r', encoding='utf-8') as f:
            original_data = json.load(f)
        
        # 解析翻译后的文本
        print("🔍 解析翻译后的文本...")
        translated_texts = parse_translated_txt(translated_txt_path)
        
        # 创建日语版本的数据
        japanese_data = []
        
        for slide in original_data:
            slide_number = slide.get('slide_number', 0)
            
            # 复制原始结构
            japanese_slide = slide.copy()
            
            # 替换文本内容
            if slide_number in translated_texts:
                translated_blocks = translated_texts[slide_number]
                
                # 更新texts字段
                if 'texts' in japanese_slide and japanese_slide['texts']:
                    for i, text_obj in enumerate(japanese_slide['texts']):
                        if i < len(translated_blocks):
                            # 保持原有的样式和位置信息，只替换content
                            text_obj['content'] = translated_blocks[i]
                        else:
                            # 如果翻译的文本块不够，保持原文
                            print(f"⚠️  第{slide_number}页文本块{i+1}没有对应的翻译，保持原文")
                
                # 如果翻译的文本块比原始的多，添加警告
                if len(translated_blocks) > len(japanese_slide.get('texts', [])):
                    print(f"⚠️  第{slide_number}页翻译的文本块数量({len(translated_blocks)})超过原始数量({len(japanese_slide.get('texts', []))})")
            
            else:
                print(f"⚠️  第{slide_number}页没有找到对应的翻译内容")
            
            japanese_data.append(japanese_slide)
        
        # 保存日语版本的JSON文件
        print("💾 保存日语版本JSON文件...")
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(japanese_data, f, ensure_ascii=False, indent=2)
        
        print(f"✅ 日语版本JSON文件生成完成！")
        print(f"📂 输入文件: {original_json_path}")
        print(f"📄 翻译文件: {translated_txt_path}")
        print(f"📋 输出文件: {output_json_path}")
        print(f"📊 共处理 {len(japanese_data)} 页幻灯片")
        
        # 统计信息
        total_translated = sum(1 for page_num in translated_texts.keys() if page_num <= len(japanese_data))
        print(f"🌐 成功翻译页面数: {total_translated}")
        
    except FileNotFoundError as e:
        print(f"❌ 文件未找到: {e}")
    except json.JSONDecodeError:
        print(f"❌ JSON文件格式错误")
    except Exception as e:
        print(f"❌ 处理过程中出现错误: {e}")

def validate_japanese_json(json_path):
    """
    验证生成的日语JSON文件
    
    Args:
        json_path (str): 日语JSON文件路径
    """
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        print(f"\n🔍 验证日语JSON文件: {json_path}")
        print(f"📊 总页面数: {len(data)}")
        
        # 检查每页的文本内容
        japanese_pages = 0
        for slide in data:
            slide_number = slide.get('slide_number', 0)
            texts = slide.get('texts', [])
            
            has_japanese = False
            for text_obj in texts:
                content = text_obj.get('content', '')
                # 简单检查是否包含日文字符
                if any('\u3040' <= char <= '\u309F' or '\u30A0' <= char <= '\u30FF' for char in content):
                    has_japanese = True
                    break
            
            if has_japanese:
                japanese_pages += 1
        
        print(f"🌐 包含日语内容的页面数: {japanese_pages}")
        print(f"📈 翻译覆盖率: {japanese_pages/len(data)*100:.1f}%")
        
    except Exception as e:
        print(f"❌ 验证过程中出现错误: {e}")

if __name__ == "__main__":
    # 设置文件路径
    original_json = "trip7_ppt_translation/extracted_content/ppt_content.json"
    translated_txt = "trip7_ppt_translation/extracted_content/ppt_texts_japanese.txt"
    output_json = "trip7_ppt_translation/extracted_content/ppt_content.japanese.json"
    
    print("🚀 开始生成日语版本PPT内容JSON文件...")
    print("=" * 50)
    
    # 检查输入文件是否存在
    if not os.path.exists(original_json):
        print(f"❌ 原始JSON文件不存在: {original_json}")
        exit(1)
    
    if not os.path.exists(translated_txt):
        print(f"❌ 翻译文本文件不存在: {translated_txt}")
        exit(1)
    
    # 执行填充操作
    fill_japanese_content(original_json, translated_txt, output_json)
    
    # 验证生成的文件
    if os.path.exists(output_json):
        validate_japanese_json(output_json)
    
    print("\n🎉 所有任务完成！")