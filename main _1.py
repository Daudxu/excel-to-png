import os
import openpyxl
import xlrd
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import time

# 获取当前脚本所在的目录，即项目的根目录
project_root = os.path.dirname(os.path.abspath(__file__))
document_directory = os.path.join(project_root, 'document_directory')  # document_directory 相对于脚本的路径
output_images_directory = os.path.join(project_root, 'output_images')  # output_images 文件夹路径

print(f"项目根目录: {project_root}")
print(f"文档目录: {document_directory}")
print(f"输出图像目录: {output_images_directory}")

# 将 Excel 文件转换为图像
def excel_to_image(excel_path, output_folder, timestamp):
    print(f"正在将 {excel_path} 转换为图像...")
    
    # 读取 Excel 文件
    if excel_path.endswith('.xlsx'):
        df = pd.read_excel(excel_path, engine='openpyxl')
    elif excel_path.endswith('.xls'):
        df = pd.read_excel(excel_path, engine='xlrd')
    else:
        print(f"不支持的文件格式: {excel_path}")
        return

    # 将 DataFrame 转换为字符串
    table_str = df.to_string(index=False)

    # 设置字体（确保你的系统中有这个字体，如果没有，请替换为系统中可用的字体）
    try:
        font = ImageFont.truetype("arial.ttf", 12)
    except IOError:
        font = ImageFont.load_default()

    # 计算图像尺寸
    lines = table_str.split('\n')
    max_width = max([font.getbbox(line)[2] for line in lines])  # 使用 getbbox() 替代 getsize()
    total_height = len(lines) * (font.getbbox(lines[0])[3] + 2)

    # 创建图像
    image = Image.new('RGB', (max_width + 20, total_height + 20), color=(255, 255, 255))
    draw = ImageDraw.Draw(image)

    # 在图像上绘制文本
    y_text = 10
    for line in lines:
        draw.text((10, y_text), line, font=font, fill=(0, 0, 0))
        y_text += font.getbbox(line)[3] + 2

    # 保存图像
    image_filename = f"{os.path.splitext(os.path.basename(excel_path))[0]}_{timestamp}.png"
    image_path = os.path.join(output_folder, image_filename)
    image.save(image_path)
    print(f"图像已保存: {image_path}")

# 查找所有 .xls 和 .xlsx 文件
for filename in os.listdir(document_directory):
    if filename.endswith(('.xls', '.xlsx')):
        print(f"找到文件: {filename}")
        excel_path = os.path.join(document_directory, filename)
        
        # 获取当前时间戳
        timestamp = int(time.time())
        
        # 生成目录名，以 Excel 文件的前缀命名，并放置在 output_images 文件夹下
        file_prefix = filename.rsplit('.', 1)[0]
        output_folder = os.path.join(output_images_directory, file_prefix)

        # 确保 output_images 目录存在
        os.makedirs(output_images_directory, exist_ok=True)

        # 确保输出文件夹存在
        os.makedirs(output_folder, exist_ok=True)
        if os.path.exists(output_folder):
            print(f"输出文件夹已创建: {output_folder}")
        else:
            print("输出文件夹创建失败")

        # 将 Excel 转换为图片
        excel_to_image(excel_path, output_folder, timestamp)
    else:
        print(f"文件 {filename} 不符合条件，跳过处理")
