import json
import os

def extract_text_from_ppt_json(json_file_path, output_txt_path):
    """
    从ppt_content.json文件中提取所有文本内容并保存到txt文件
    
    Args:
        json_file_path (str): ppt_content.json文件路径
        output_txt_path (str): 输出的txt文件路径
    """
    try:
        # 读取JSON文件
        with open(json_file_path, 'r', encoding='utf-8') as f:
            slides_data = json.load(f)
        
        # 提取文本内容
        extracted_texts = []
        
        for slide in slides_data:
            slide_number = slide.get('slide_number', 0)
            extracted_texts.append(f"\n=== 第{slide_number}页 ===")
            
            # 提取该页面的所有文本
            texts = slide.get('texts', [])
            if texts:
                for i, text_obj in enumerate(texts, 1):
                    content = text_obj.get('content', '').strip()
                    if content:
                        extracted_texts.append(f"\n文本块 {i}:")
                        extracted_texts.append(content)
            else:
                extracted_texts.append("(该页面无文本内容)")
        
        # 写入txt文件
        with open(output_txt_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(extracted_texts))
        
        print(f"✅ 文本提取完成！")
        print(f"📂 输入文件: {json_file_path}")
        print(f"📄 输出文件: {output_txt_path}")
        print(f"📊 共处理 {len(slides_data)} 页幻灯片")
        
        # 统计信息
        total_text_blocks = sum(len(slide.get('texts', [])) for slide in slides_data)
        print(f"📝 总文本块数: {total_text_blocks}")
        
    except FileNotFoundError:
        print(f"❌ 错误: 找不到文件 {json_file_path}")
    except json.JSONDecodeError:
        print(f"❌ 错误: JSON文件格式不正确")
    except Exception as e:
        print(f"❌ 处理过程中出现错误: {e}")

def extract_simple_text(json_file_path, output_txt_path):
    """
    简化版本：只提取纯文本内容，不包含页面信息
    
    Args:
        json_file_path (str): ppt_content.json文件路径
        output_txt_path (str): 输出的txt文件路径
    """
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            slides_data = json.load(f)
        
        all_texts = []
        
        for slide in slides_data:
            texts = slide.get('texts', [])
            for text_obj in texts:
                content = text_obj.get('content', '').strip()
                if content:
                    all_texts.append(content)
        
        # 写入txt文件
        with open(output_txt_path, 'w', encoding='utf-8') as f:
            f.write("\n\n".join(all_texts))
        
        print(f"✅ 简化文本提取完成！共提取 {len(all_texts)} 个文本块")
        
    except Exception as e:
        print(f"❌ 处理过程中出现错误: {e}")

if __name__ == "__main__":
    # 设置文件路径
    json_file = "trip7_ppt_translation/extracted_content/ppt_content.json"
    
    # 详细版本输出
    detailed_output = "trip7_ppt_translation/extracted_content/ppt_texts_detailed.txt"
    extract_text_from_ppt_json(json_file, detailed_output)
    
    # 简化版本输出
    simple_output = "trip7_ppt_translation/extracted_content/ppt_texts_simple.txt"
    extract_simple_text(json_file, simple_output)
    
    print("\n🎉 所有文本提取任务完成！")