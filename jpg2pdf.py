from PIL import Image
import os


def merge_images_to_pdf(input_folder, output_pdf):
    """
    将指定路径下的所有 JPG 图像合并为一个 PDF 文件。

    :param input_folder: 包含 JPG 图像的文件夹路径
    :param output_pdf: 输出 PDF 文件路径
    """
    # 获取文件夹中所有 .jpg 文件并按文件名排序
    image_files = sorted(
        [f for f in os.listdir(input_folder) if f.lower().endswith('.jpg')],
        key=lambda x: x.lower()
    )

    # 检查是否有图像文件
    if not image_files:
        print("未找到任何 JPG 文件！")
        return

    images = []
    for image_file in image_files:
        print(image_file)
        # 拼接完整路径并打开图像
        img_path = os.path.join(input_folder, image_file)
        img = Image.open(img_path).convert('RGB')  # 转为 RGB 模式
        images.append(img)

    # 将所有图像合并到一个 PDF 文件中
    try:
        # 第一张图像作为基准，其余作为附加页面
        images[0].save(output_pdf, save_all=True, append_images=images[1:])
        print(f"PDF 文件已成功保存到 {output_pdf}")
    except Exception as e:
        print(f"生成 PDF 文件失败：{e}")


# 示例用法
input_folder = r"C:\Users\Alan\Downloads\新建文件夹"  # 替换为实际的图像文件夹路径
output_pdf = r"C:\Users\Alan\Downloads\新建文件夹\发票.pdf"  # 输出 PDF 文件路径
merge_images_to_pdf(input_folder, output_pdf)
