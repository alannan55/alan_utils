import glob
import os
import shutil
import cv2
import numpy as np


def cv_imread(file_path):
    cv_img = cv2.imdecode(np.fromfile(file_path,dtype=np.uint8),-1)
    return cv_img


def cp_files(src_root, tgt_root):

    for file_path in glob.glob(src_root + "\*"):
        img_array = cv_imread(file_path)
        h, w, _ = img_array.shape
        if w >= 3840 and w >= h:
            tgt_path = file_path.replace(src_root, tgt_root)
            shutil.copy(file_path, tgt_path)


if __name__ == '__main__':
    src_root = r"H:\图片\壁纸\次元壁纸"
    tgt_root = r"H:\图片\壁纸\新建文件夹"
    cp_files(src_root, tgt_root)