from pptx import Presentation
import os
import json

def extract_ppt_content(ppt_path, output_dir="ppt_output", save_images=True):
    """
    提取PPTX文件的纯文本内容，按页数一一对应
    
    Args:
        ppt_path (str): PPT文件路径
        output_dir (str): 输出目录
        save_images (bool): 是否保存图片文件
    
    Returns:
        list: 包含所有幻灯片信息的列表
    """
    # 检查文件是否存在
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PPT文件不存在: {ppt_path}")
    
    prs = Presentation(ppt_path)
    os.makedirs(output_dir, exist_ok=True)
    
    all_slides = []
    
    for i, slide in enumerate(prs.slides, start=1):
        slide_info = {
            "slide_number": i,
            "texts": [],
            "images": []
        }
        
        for shape in slide.shapes:
            # 提取纯文字内容
            if hasattr(shape, "text") and shape.text.strip():
                slide_info["texts"].append({
                    "content": shape.text.strip()
                })
            
            # 提取图片（如果需要）
            if shape.shape_type == 13:  # 13表示图片类型
                try:
                    image = shape.image
                    image_bytes = image.blob
                    image_ext = image.ext
                    image_name = f"slide_{i}_img_{len(slide_info['images']) + 1}.{image_ext}"
                    image_path = os.path.join(output_dir, image_name)
                    
                    if save_images:
                        with open(image_path, "wb") as f:
                            f.write(image_bytes)
                    
                    slide_info["images"].append({
                        "filename": image_name
                    })
                except Exception as e:
                    print(f"⚠️ 提取第{i}页图片时出错: {e}")
        
        all_slides.append(slide_info)
    
    # 导出为JSON
    json_path = os.path.join(output_dir, "ppt_content.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_slides, f, ensure_ascii=False, indent=2)
    
    print(f"✅ 提取完成，共 {len(all_slides)} 页。结果保存在：{json_path}")
    
    # 打印摘要信息
    total_texts = sum(len(slide["texts"]) for slide in all_slides)
    total_images = sum(len(slide["images"]) for slide in all_slides)
    print(f"📊 统计信息: 文本块 {total_texts} 个，图片 {total_images} 张")
    
    return all_slides

def print_content_summary(slides_data):
    """
    打印内容摘要
    """
    print("\n📋 内容摘要:")
    for slide in slides_data:
        slide_num = slide["slide_number"]
        text_count = len(slide["texts"])
        image_count = len(slide["images"])
        print(f"第{slide_num}页: {text_count}个文本块, {image_count}张图片")
        
        # 显示前50个字符的文本预览
        if slide["texts"]:
            preview = slide["texts"][0]["content"][:50] + "..." if len(slide["texts"][0]["content"]) > 50 else slide["texts"][0]["content"]
            print(f"  文本预览: {preview}")

# 示例用法
if __name__ == "__main__":
    # 从trip7_ppt_translation/chinese目录读取PPT文件
    chinese_dir = "trip7_ppt_translation/chinese"
    ppt_files = [f for f in os.listdir(chinese_dir) if f.endswith('.pptx')]
    
    if not ppt_files:
        print("❌ 在chinese目录中未找到PPTX文件")
        exit(1)
    
    # 处理第一个找到的PPTX文件
    ppt_file = os.path.join(chinese_dir, ppt_files[0])
    print(f"📂 处理文件: {ppt_file}")
    
    try:
        # 提取内容到trip7_ppt_translation/extracted_content目录
        slides_data = extract_ppt_content(ppt_file, output_dir="trip7_ppt_translation/extracted_content")
        
        # 打印摘要
        print_content_summary(slides_data)
        
    except FileNotFoundError as e:
        print(f"❌ 错误: {e}")
        print("请确保PPT文件存在于chinese目录中")
    except Exception as e:
        print(f"❌ 处理过程中出现错误: {e}")