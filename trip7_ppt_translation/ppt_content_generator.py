from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import json
import os
from PIL import Image

def create_ppt_from_json(json_path, output_ppt_path, template_ppt_path=None):
    """
    从日语JSON文件生成PPT文件
    
    Args:
        json_path (str): 日语JSON文件路径
        output_ppt_path (str): 输出PPT文件路径
        template_ppt_path (str): 模板PPT文件路径（可选）
    """
    try:
        # 读取JSON数据
        print("📖 读取日语JSON文件...")
        with open(json_path, 'r', encoding='utf-8') as f:
            slides_data = json.load(f)
        
        # 创建新的演示文稿或使用模板
        if template_ppt_path and os.path.exists(template_ppt_path):
            print(f"📋 使用模板文件: {template_ppt_path}")
            prs = Presentation(template_ppt_path)
            # 清空现有幻灯片
            slide_ids = [slide.slide_id for slide in prs.slides]
            for slide_id in slide_ids:
                rId = prs.slides._sldIdLst[0].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[0]
        else:
            print("📋 创建新的演示文稿")
            prs = Presentation()
            # 删除默认的空白幻灯片
            if len(prs.slides) > 0:
                rId = prs.slides._sldIdLst[0].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[0]
        
        # 设置幻灯片尺寸（如果JSON中有尺寸信息）
        if slides_data and 'slide_width' in slides_data[0]:
            prs.slide_width = slides_data[0]['slide_width']
            prs.slide_height = slides_data[0]['slide_height']
        
        print(f"🔄 开始生成 {len(slides_data)} 页幻灯片...")
        
        # 为每个幻灯片数据创建幻灯片
        for slide_data in slides_data:
            slide_number = slide_data.get('slide_number', 1)
            print(f"📄 生成第 {slide_number} 页...")
            
            # 添加空白幻灯片
            blank_slide_layout = prs.slide_layouts[6]  # 空白布局
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # 添加图片
            images_dir = os.path.dirname(json_path)
            for img_info in slide_data.get('images', []):
                add_image_to_slide(slide, img_info, images_dir)
            
            # 添加文本
            for text_info in slide_data.get('texts', []):
                add_text_to_slide(slide, text_info)
        
        # 保存PPT文件
        print(f"💾 保存PPT文件: {output_ppt_path}")
        prs.save(output_ppt_path)
        
        print(f"✅ 日语PPT文件生成完成！")
        print(f"📂 输入文件: {json_path}")
        print(f"📋 输出文件: {output_ppt_path}")
        print(f"📊 共生成 {len(slides_data)} 页幻灯片")
        
    except FileNotFoundError as e:
        print(f"❌ 文件未找到: {e}")
    except json.JSONDecodeError:
        print(f"❌ JSON文件格式错误")
    except Exception as e:
        print(f"❌ 生成PPT过程中出现错误: {e}")

def add_image_to_slide(slide, img_info, images_dir):
    """
    向幻灯片添加图片
    
    Args:
        slide: 幻灯片对象
        img_info (dict): 图片信息
        images_dir (str): 图片目录路径
    """
    try:
        filename = img_info.get('filename')
        if not filename:
            return
        
        image_path = os.path.join(images_dir, filename)
        if not os.path.exists(image_path):
            print(f"⚠️  图片文件不存在: {image_path}")
            return
        
        # 获取位置信息
        position = img_info.get('position', {})
        left = position.get('left', 0)
        top = position.get('top', 0)
        width = position.get('width', Inches(2))
        height = position.get('height', Inches(2))
        
        # 添加图片到幻灯片
        slide.shapes.add_picture(
            image_path,
            left=left,
            top=top,
            width=width,
            height=height
        )
        
    except Exception as e:
        print(f"⚠️  添加图片时出错: {e}")

def add_text_to_slide(slide, text_info):
    """
    向幻灯片添加文本
    
    Args:
        slide: 幻灯片对象
        text_info (dict): 文本信息
    """
    try:
        content = text_info.get('content', '')
        if not content.strip():
            return
        
        # 获取位置信息
        position = text_info.get('position', {})
        left = position.get('left', 0)
        top = position.get('top', 0)
        width = position.get('width', Inches(3))
        height = position.get('height', Inches(1))
        
        # 添加文本框
        textbox = slide.shapes.add_textbox(
            left=left,
            top=top,
            width=width,
            height=height
        )
        
        text_frame = textbox.text_frame
        text_frame.clear()  # 清空默认段落
        
        # 应用样式信息
        style_info = text_info.get('style', {})
        paragraphs_info = style_info.get('paragraphs', [])
        
        if paragraphs_info:
            # 使用详细的样式信息
            for para_info in paragraphs_info:
                if text_frame.paragraphs:
                    p = text_frame.paragraphs[0] if len(text_frame.paragraphs) == 1 else text_frame.add_paragraph()
                else:
                    p = text_frame.add_paragraph()
                
                # 设置段落对齐方式
                alignment = para_info.get('alignment')
                if alignment:
                    try:
                        if 'CENTER' in str(alignment).upper():
                            p.alignment = PP_ALIGN.CENTER
                        elif 'RIGHT' in str(alignment).upper():
                            p.alignment = PP_ALIGN.RIGHT
                        elif 'LEFT' in str(alignment).upper():
                            p.alignment = PP_ALIGN.LEFT
                    except:
                        pass
                
                # 添加文本运行
                runs_info = para_info.get('runs', [])
                if runs_info:
                    for run_info in runs_info:
                        run_text = run_info.get('text', '')
                        if run_text:
                            run = p.add_run()
                            run.text = run_text
                            
                            # 应用字体样式
                            apply_font_style(run, run_info)
                else:
                    # 如果没有runs信息，使用段落文本
                    para_text = para_info.get('text', '')
                    if para_text:
                        run = p.add_run()
                        run.text = para_text
        else:
            # 简单文本，没有详细样式信息
            p = text_frame.paragraphs[0]
            run = p.add_run()
            run.text = content
            
            # 设置默认日语字体
            if run.font:
                run.font.name = 'Yu Gothic'  # 日语常用字体
                run.font.size = Pt(18)
        
    except Exception as e:
        print(f"⚠️  添加文本时出错: {e}")

def apply_font_style(run, run_info):
    """
    应用字体样式
    
    Args:
        run: 文本运行对象
        run_info (dict): 运行样式信息
    """
    try:
        font = run.font
        
        # 字体名称
        font_name = run_info.get('font_name')
        if font_name:
            font.name = font_name
        else:
            # 默认日语字体
            font.name = 'Yu Gothic'
        
        # 字体大小
        font_size = run_info.get('font_size')
        if font_size:
            font.size = Pt(font_size)
        
        # 粗体
        bold = run_info.get('bold')
        if bold is not None:
            font.bold = bold
        
        # 斜体
        italic = run_info.get('italic')
        if italic is not None:
            font.italic = italic
        
        # 下划线
        underline = run_info.get('underline')
        if underline and underline != 'None':
            font.underline = True
        
        # 字体颜色
        color = run_info.get('color')
        if color and color != 'None':
            try:
                # 解析RGB颜色
                if color.startswith('RGBColor'):
                    # 提取RGB值
                    rgb_match = color.replace('RGBColor(', '').replace(')', '')
                    if ',' in rgb_match:
                        r, g, b = map(int, rgb_match.split(','))
                        font.color.rgb = RGBColor(r, g, b)
            except:
                pass
        
    except Exception as e:
        print(f"⚠️  应用字体样式时出错: {e}")

def validate_generated_ppt(ppt_path):
    """
    验证生成的PPT文件
    
    Args:
        ppt_path (str): PPT文件路径
    """
    try:
        prs = Presentation(ppt_path)
        
        print(f"\n🔍 验证生成的PPT文件: {ppt_path}")
        print(f"📊 总幻灯片数: {len(prs.slides)}")
        print(f"📐 幻灯片尺寸: {prs.slide_width} x {prs.slide_height}")
        
        # 检查每页内容
        total_shapes = 0
        for i, slide in enumerate(prs.slides, 1):
            shape_count = len(slide.shapes)
            total_shapes += shape_count
            print(f"第{i}页: {shape_count}个形状")
        
        print(f"📈 总形状数: {total_shapes}")
        print("✅ PPT文件验证完成")
        
    except Exception as e:
        print(f"❌ 验证PPT文件时出错: {e}")

if __name__ == "__main__":
    # 设置文件路径
    japanese_json = "trip7_ppt_translation/extracted_content/ppt_content.japanese.json"
    output_ppt = "trip7_ppt_translation/japanese/Web3与元宇宙_日语版.pptx"
    
    print("🚀 开始生成日语PPT文件...")
    print("=" * 50)
    
    # 检查输入文件是否存在
    if not os.path.exists(japanese_json):
        print(f"❌ 日语JSON文件不存在: {japanese_json}")
        print("请先运行ppt_content_filler.py生成日语JSON文件")
        exit(1)
    
    # 创建输出目录
    os.makedirs(os.path.dirname(output_ppt), exist_ok=True)
    
    # 生成PPT文件
    create_ppt_from_json(japanese_json, output_ppt)
    
    # 验证生成的文件
    if os.path.exists(output_ppt):
        validate_generated_ppt(output_ppt)
    
    print("\n🎉 日语PPT生成任务完成！")