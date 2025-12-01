from PIL import Image
import os
import glob

def webp_to_jpg(img):
    tmp_file = img.replace('webp', 'png')
    # 临时文件的文件名
    to_file = img.replace('webp', 'jpg')
    # 转换后的文件名

    tmp_image = Image.open(img)
    tmp_image.save(tmp_file, 'PNG')
    # 创建临时文件并保存

    try:
        to_img = Image.open(tmp_file).convert('RGB')
        to_img.save(to_file, 'JPEG', quality=100)
        # 创建jpg文件并保存
        os.remove(img)
        # 删除原始文件
    except:
        print(img + '转换错误')
    finally:
        if os.path.exists(tmp_file):
            os.remove(tmp_file)
            # 删除临时文件


if __name__ == '__main__':
    img_root = r"C:\Users\alan\Desktop\su7痛车"
    for webp_path in glob.glob(os.path.join(img_root, "*.webp")):
        webp_to_jpg(webp_path)
