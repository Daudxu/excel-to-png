import openpyxl
from PIL import Image, ImageDraw, ImageFont

def excel_to_image(excel_path, output_image_path):
    # 载入 Excel 文件
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active

    # 获取 Excel 数据
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)

    # 设置图像大小
    cell_width = 100
    cell_height = 30
    img_width = cell_width * len(data[0])
    img_height = cell_height * len(data)
    
    # 创建图像
    image = Image.new('RGB', (img_width, img_height), 'white')
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()

    # 绘制单元格内容
    for i, row in enumerate(data):
        for j, value in enumerate(row):
            x = j * cell_width
            y = i * cell_height
            draw.rectangle([x, y, x + cell_width, y + cell_height], outline='black')
            draw.text((x + 5, y + 5), str(value), font=font, fill='black')

    # 保存图像
    image.save(output_image_path)

# 示例使用
excel_path = 'd:\\testWork\\ziyuan\\excel-to-png\\document_directory\\员工季度业绩排名表.xlsx'
output_image_path = 'd:\\testWork\\ziyuan\\excel-to-png\\output_images\\员工季度业绩排名表.png'
excel_to_image(excel_path, output_image_path)
