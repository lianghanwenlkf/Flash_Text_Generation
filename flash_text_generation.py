import os
import random
import cv2
import fitz  # PyMuPDF
import numpy as np
import pandas as pd
from time import sleep
from datetime import datetime
from tqdm import tqdm
from PIL import Image


def select_random_images(folder_list, output_folder, log_file_path, num_images=20,
                         limit_width=170, limit_height=139):
    select_random_images_list = []
    # 循环遍历给定的文件夹列表
    for i, folder in enumerate(folder_list):
        image_files = [f for f in os.listdir(folder) if f.endswith('.jpg') or f.endswith('.png')]
        if len(image_files) < num_images:
            with open(log_file_path, 'a', encoding='utf-8') as file:
                file.write(f'{folder} < {num_images}\n')
                print(f'{folder} < {num_images}')
        # 从当前文件夹中随机选择num_images张图片
        selected_images = random.sample(image_files, len(image_files))
        num_now = 0
        # 处理每张选中的图片
        for j, image_name in enumerate(selected_images):
            if num_now >= num_images:
                break
            image_path = os.path.join(folder, image_name)
            image = Image.open(image_path)
            width, height = image.size

            # 检查分辨率是否在指定范围内
            if width == limit_width and height == limit_height:
                # 将图片拷贝到输出文件夹
                image.save(os.path.join(output_folder, image_name))
                select_random_images_list.append(os.path.join(output_folder, image_name))

                num_now += 1
    return select_random_images_list


def images_to_video(select_random_images_list, output_video_path, fps, width, height):
    # 创建视频写入器
    fourcc = cv2.VideoWriter_fourcc(*'mp4v')  # 可以根据需要更改编解码器
    video_writer = cv2.VideoWriter(output_video_path, fourcc, fps, (width, height))

    for image_file in select_random_images_list:
        img = cv2.imdecode(np.fromfile(image_file, dtype=np.uint8), -1)
        video_writer.write(img)

    video_writer.release()


def create_folder(folder_path):
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        # 如果不存在，则创建文件夹
        os.makedirs(folder_path)
        print(f"文件夹创建成功: {folder_path}")
    else:
        print(f"文件夹已存在: {folder_path}")


def find_and_capture_character_in_pdfs(index, character, pdf_directory, save_directory,
                                       x_crop_factor, y_crop_factor, width, height):
    # 创建保存截图的目录
    if not os.path.exists(save_directory):
        os.makedirs(save_directory)

    # 遍历指定目录下的所有PDF文件
    for pdf_file in os.listdir(pdf_directory):
        if pdf_file.endswith('.pdf'):
            pdf_path = os.path.join(pdf_directory, pdf_file)
            pdf_document = fitz.open(pdf_path)
            # 遍历PDF的所有页面
            for page_number in range(pdf_document.page_count):
                page = pdf_document[page_number]

                # 获取页面文本
                text = page.get_text("text")

                # 检查页面是否包含指定的中文字符
                if character in text:
                    # 获取字符在页面上的矩形边界框
                    rect = page.search_for(character)[0]

                    # 获取页面尺寸
                    page_width, page_height = page.rect.width, page.rect.height

                    # 设置截图区域
                    x_min, y_min, x_max, y_max = (max(0, rect.x0 - page_width / x_crop_factor),
                                                  max(0, rect.y0 - page_height / y_crop_factor),
                                                  min(page_width, rect.x1 + page_width / x_crop_factor),
                                                  min(page_height, rect.y1 + page_height / y_crop_factor))

                    if width-2 <= (x_max-x_min)*2 <= width+2 and height-2 <= (y_max-y_min)*2 <= height+2:
                        # 渲染页面为图像
                        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=fitz.Rect(x_min, y_min, x_max, y_max))

                        # 将图像数据转换为PIL Image
                        pil_image = Image.frombuffer("RGB", (pixmap.width, pixmap.height), pixmap.samples)

                        # 保存以字符为中心的页面截图
                        save_path = os.path.join(save_directory,
                                                 f"{index}_{character}_{pdf_file[:-4]}_page{page_number + 1}.png")
                        pil_image.save(save_path)

            # 关闭PDF文件
            pdf_document.close()


def read_excel_column(excel_path, sheet_name='Sheet1', column_index=0):
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path, sheet_name=sheet_name)

        # 获取指定列的数据并转换为列表
        column_data = df.iloc[:, column_index].tolist()

        return column_data

    except Exception as e:
        print(f"Error: {e}")
        return None


def main():
    # 输入路径
    pdf_folder_path = './1_pdf_files'  # 存放PDF文件的目录
    name_file_path = './2_names/names.xlsx'  # 存放好友列表

    # 输出路径
    character_images_folder_path = './3_character_images'  # 存放字符截图
    selected_images_folder_path = './4_selected_images'  # 存放分辨率符合要求的截图
    output_video_folder_path = './5_output_video'  # 存放输出的视频
    log_path = './6_log'  # 存放日志

    # 截图参数
    num_images = 20  # 每个字截取多少张图片
    x_crop_factor = 12  # 横向裁剪倍率，根据pdf页面大小和字体大小调整。若使用提供的素材，则保持默认即可
    y_crop_factor = 24  # 纵向裁剪倍率，根据pdf页面大小和字体大小调整。若使用提供的素材，则保持默认即可

    # 视频参数
    fps = 20  # 每秒播放多少张图片
    width = 170  # 每张图片的宽度，因为截图的分辨率有小幅度波动，选择平均值填写。若使用提供的素材，则保持默认即可
    height = 139  # 每张图片的高度，因为截图的分辨率有小幅度波动，选择平均值填写。若使用提供的素材，则保持默认即可

    # 设置祝福语，sentence = sentence_1 + name + sentence_2 + selected_wish + sentence_3
    sentence_1 = '恭祝'
    sentence_2 = '新年快乐龙年大吉'
    sentence_3 = '期待下次见面梁某人敬上二零二四年除夕'
    # 随机祝福语库
    wishes = [
        "龙马精神",
        "一帆风顺",
        "步步高升",
        "福星高照",
        "龙飞凤舞",
        "事业有成",
        "笑口常开",
        "身体健康",
        "幸福美满",
        "心想事成",
        "吉祥如意",
        "龙送吉祥",
        "龙瑞盈门",
        "龙来运转",
        "龙跃新程",
        "龙跃千里前程锦绣",
        "龙年发财喜事连连",
        "龙舞云川喜气洋洋",
        "龙舞九天好运连连",
        "舞龙迎春好事如潮",
        "龙舞云端喜乐无边",
        "龙凤呈祥幸福美满",
        "龙腾四海",
        "龙翔虎跃"
    ]

    create_folder(selected_images_folder_path)
    create_folder(output_video_folder_path)
    create_folder(log_path)
    sleep(0.1)
    log_file_path = log_path + '/' + datetime.now().strftime("%Y-%m-%d %H-%M-%S") + '.txt'

    name_list = read_excel_column(name_file_path)
    for name in tqdm(name_list):
        sleep(0.1)
        print(name)
        selected_wish = random.choice(wishes)
        sentence = sentence_1 + name + sentence_2 + selected_wish + sentence_3

        folder_list = []
        for index, character_to_find in enumerate(sentence):
            save_folder_path = f'{character_images_folder_path}/{character_to_find}'  # 保存截图的目录
            folder_list.append(save_folder_path)
            if not os.path.exists(save_folder_path):
                os.makedirs(save_folder_path)
                print(f"文件夹创建成功: {save_folder_path}")
                find_and_capture_character_in_pdfs(index, character_to_find, pdf_folder_path, save_folder_path,
                                                   x_crop_factor, y_crop_factor, width, height)
            else:
                print(f"文件夹已存在: {save_folder_path}")

        # 选择随机图片
        personal_selected_images_folder_path = selected_images_folder_path + '/' + name  # 输出文件夹
        create_folder(personal_selected_images_folder_path)
        select_random_images_list = select_random_images(folder_list, personal_selected_images_folder_path,
                                                         log_file_path, num_images, width, height)

        # 创建视频
        output_video_path = output_video_folder_path + '/' + name + '祝福.mp4'  # 输出视频文件名
        images_to_video(select_random_images_list, output_video_path, fps, width, height)


if __name__ == "__main__":
    main()
