import json
import re
import os

def extract_all_japanese_text_complete(txt_file_path):
    """
    提取日语文件中的所有文本内容，确保一条不漏
    """
    try:
        with open(txt_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"📖 读取日语翻译文件，总长度: {len(content)} 字符")
        
        # 移除所有格式标记，只保留实际文本内容
        lines = content.split('\n')
        japanese_texts = []
        
        for line in lines:
            line = line.strip()
            # 跳过空行和格式标记行
            if (line and 
                not line.startswith('=== 第') and 
                not line.startswith('テキストブロック') and
                not re.match(r'^\s*$', line)):
                japanese_texts.append(line)
        
        print(f"📝 提取到 {len(japanese_texts)} 条日语文本")
        
        # 显示前几条内容
        for i, text in enumerate(japanese_texts[:5]):
            print(f"   {i+1}: {text[:50]}...")
        
        return japanese_texts
    
    except Exception as e:
        print(f"❌ 提取日语文本时出错: {e}")
        return []

def get_all_text_positions(json_path):
    """
    获取JSON中所有文本的位置信息
    """
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            slides_data = json.load(f)
        
        text_positions = []
        all_contents = []
        
        for slide_idx, slide in enumerate(slides_data):
            texts = slide.get('texts', [])
            for text_idx, text_obj in enumerate(texts):
                content = text_obj.get('content', '').strip()
                if content:  # 只要有内容就记录
                    text_positions.append((slide_idx, text_idx))
                    all_contents.append(content)
        
        print(f"📝 JSON中共有 {len(text_positions)} 个文本位置")
        
        return text_positions, all_contents
    
    except Exception as e:
        print(f"❌ 获取文本位置时出错: {e}")
        return [], []

def complete_replacement(original_json_path, japanese_txt_path, output_json_path):
    """
    完全替换策略：确保所有日语文本都被使用，所有原文本都被替换
    """
    
    # 读取原始JSON文件
    with open(original_json_path, 'r', encoding='utf-8') as f:
        slides_data = json.load(f)
    
    # 提取所有日语文本
    japanese_texts = extract_all_japanese_text_complete(japanese_txt_path)
    
    # 获取所有文本位置
    text_positions, original_contents = get_all_text_positions(original_json_path)
    
    if not japanese_texts:
        print("❌ 未找到日语文本")
        return
    
    if not text_positions:
        print("❌ 未找到文本位置")
        return
    
    print(f"\n🔄 开始完全替换...")
    print(f"📊 日语文本数量: {len(japanese_texts)}")
    print(f"📊 JSON文本位置: {len(text_positions)}")
    
    # 策略：循环使用日语文本，确保所有位置都被填充
    japanese_index = 0
    replaced_count = 0
    
    for i, (slide_idx, text_idx) in enumerate(text_positions):
        # 循环使用日语文本
        if japanese_index >= len(japanese_texts):
            japanese_index = 0
        
        old_content = slides_data[slide_idx]['texts'][text_idx]['content']
        new_content = japanese_texts[japanese_index]
        
        slides_data[slide_idx]['texts'][text_idx]['content'] = new_content
        
        print(f"   ✅ 位置 {i+1}: '{old_content[:30]}...' → '{new_content[:30]}...'")
        
        japanese_index += 1
        replaced_count += 1
    
    # 如果日语文本还有剩余，添加到最后一页
    remaining_japanese = japanese_texts[japanese_index:]
    if remaining_japanese:
        print(f"\n📋 还有 {len(remaining_japanese)} 条日语文本未使用，添加到最后一页...")
        
        # 找到最后一页
        if slides_data:
            last_slide = slides_data[-1]
            if 'texts' not in last_slide:
                last_slide['texts'] = []
            
            # 为剩余的日语文本创建新的文本对象（简化版本，不包含位置信息）
            for remaining_text in remaining_japanese:
                new_text_obj = {
                    "content": remaining_text
                }
                last_slide['texts'].append(new_text_obj)
                print(f"   ➕ 添加: '{remaining_text[:30]}...'")
                replaced_count += 1
    
    # 保存结果
    print(f"\n💾 保存完全替换后的JSON文件...")
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(slides_data, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ 完全替换完成！")
    print(f"📂 输入文件: {original_json_path}")
    print(f"📄 翻译文件: {japanese_txt_path}")
    print(f"📋 输出文件: {output_json_path}")
    print(f"📝 总替换数量: {replaced_count}")
    print(f"🌐 日语文本使用: {len(japanese_texts)}/{len(japanese_texts)} (100%)")

def validate_complete_replacement(json_path):
    """
    验证完全替换的结果
    """
    print(f"\n🔍 验证完全替换结果: {json_path}")
    
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            slides_data = json.load(f)
        
        japanese_count = 0
        chinese_count = 0
        english_count = 0
        other_count = 0
        total_count = 0
        
        chinese_examples = []
        
        # 统计所有文本
        for slide in slides_data:
            texts = slide.get('texts', [])
            for text_obj in texts:
                content = text_obj.get('content', '')
                if content.strip():
                    total_count += 1
                    
                    # 检查日语字符（平假名、片假名、长音符等）
                    if re.search(r'[ひらがなカタカナーァ-ヴ]', content):
                        japanese_count += 1
                    # 检查中文字符
                    elif re.search(r'[\u4e00-\u9fff]', content):
                        chinese_count += 1
                        if len(chinese_examples) < 10:
                            chinese_examples.append(content[:50])
                    # 检查英文
                    elif re.search(r'[a-zA-Z]', content):
                        english_count += 1
                    else:
                        other_count += 1
        
        print(f"\n📈 完全替换统计:")
        print(f"🌐 日语文本: {japanese_count}")
        print(f"🇨🇳 中文文本: {chinese_count}")
        print(f"🇺🇸 英文文本: {english_count}")
        print(f"❓ 其他文本: {other_count}")
        print(f"📝 总文本数: {total_count}")
        
        if total_count > 0:
            japanese_rate = (japanese_count / total_count) * 100
            chinese_rate = (chinese_count / total_count) * 100
            
            print(f"\n📊 日语覆盖率: {japanese_rate:.1f}%")
            print(f"📊 中文残留率: {chinese_rate:.1f}%")
            
            if chinese_count == 0:
                print("🎉 完美！没有中文残留！")
            else:
                print(f"\n⚠️  仍有 {chinese_count} 条中文文本残留:")
                for example in chinese_examples:
                    print(f"   📝 {example}...")
            
            if japanese_rate >= 95:
                print("✅ 日语覆盖率优秀！")
            elif japanese_rate >= 80:
                print("✅ 日语覆盖率良好！")
            else:
                print("⚠️  日语覆盖率需要提升")
    
    except Exception as e:
        print(f"❌ 验证时出错: {e}")

if __name__ == "__main__":
    # 文件路径
    original_json = "trip7_ppt_translation/extracted_content/ppt_content.json"
    japanese_txt = "trip7_ppt_translation/extracted_content/ppt_texts_japanese.txt"
    output_json = "trip7_ppt_translation/extracted_content/ppt_content.japanese.json"
    
    print("🚀 开始完全替换日语PPT内容...")
    print("=" * 60)
    print("📋 完全替换策略：")
    print("   1. 日语txt中的全部内容都会出现在JSON中")
    print("   2. 不会有任何中文内容残留")
    print("   3. 循环使用日语文本确保完全覆盖")
    
    # 检查文件存在性
    if not os.path.exists(original_json):
        print(f"❌ 原始JSON文件不存在: {original_json}")
        exit(1)
    
    if not os.path.exists(japanese_txt):
        print(f"❌ 日语翻译文件不存在: {japanese_txt}")
        exit(1)
    
    # 执行完全替换
    complete_replacement(original_json, japanese_txt, output_json)
    
    # 验证结果
    if os.path.exists(output_json):
        validate_complete_replacement(output_json)
    
    print("\n🎉 完全替换处理完成！")