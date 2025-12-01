from PIL import Image
from PyPDF2 import PdfMerger
import os

def images_to_pdf(input_folder, output_pdf):
    # 临时PDF文件列表
    temp_pdf_list = []

    # 将每张图片转换为单页PDF并保存到临时文件
    for i in range(1, 11):  # Adjust range as needed
        img_path = os.path.join(input_folder, f"{i}.jpg")
        img = Image.open(img_path)
        temp_pdf_path = os.path.join(input_folder, f"{i}.pdf")
        img.convert('RGB').save(temp_pdf_path)
        temp_pdf_list.append(temp_pdf_path)

    # 合并所有临时PDF文件
    merger = PdfMerger()
    for pdf in temp_pdf_list:
        merger.append(pdf)

    # 写入输出PDF文件
    with open(output_pdf, "wb") as f:
        merger.write(f)

    # 删除临时PDF文件
    for pdf in temp_pdf_list:
        os.remove(pdf)

if __name__ == "__main__":
    input_folder = "C:\zhaonan\找工作准备\华为OD\OD隐私保护声明"  # Replace with your image folder path
    output_pdf = "output.pdf"
    images_to_pdf(input_folder, output_pdf)
