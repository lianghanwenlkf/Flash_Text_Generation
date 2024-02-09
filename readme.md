# 快闪文字生成
## 1 素材准备
### 1.1 书籍pdf
#### 路径：./pdf_files
#### 文件名：{书名}.pdf
#### 内容：中文书籍pdf，需要可搜索文字，不能是图片或者扫描件，可直接使用提供的素材
### 2.2 好友列表
#### 路径：./<br>
#### 文件名：name.xlsx<br>
#### 内容：在第一列打出所有想要发送祝福的好友的名字，可通过QQ邮箱或者留痕（MemoTrace）软件导出联系人
## 2 参数设置
### 2.1 输入路径
#### （1）pdf_folder_path  # 书籍pdf
#### （2）name_file_path  # 好友列表
### 2.2 输出路径
#### （1）character_images_folder_path  # 存放字符截图
#### （2）selected_images_folder_path  # 存放分辨率符合要求的截图
#### （3）output_video_folder_path  # 存放输出视频
#### （4）log_path  # 存放运行日志
### 2.3 截图参数
#### （1）num_images = 20  # 每个字截取多少张图片
#### （2）x_crop_factor = 12  # 横向裁剪倍率，根据pdf页面大小和字体大小调整。若使用提供的素材，则保持默认即可
#### （3）y_crop_factor = 24  # 纵向裁剪倍率，根据pdf页面大小和字体大小调整。若使用提供的素材，则保持默认即可

### 2.4 视频参数
#### （1）fps = 20  # 每秒播放多少张图片
#### （2）width = 170  # 每张图片的宽度，因为截图的分辨率有小幅度波动，选择平均值填写。若使用提供的素材，则保持默认即可
#### （3）height = 139  # 每张图片的高度，因为截图的分辨率有小幅度波动，选择平均值填写。若使用提供的素材，则保持默认即可

### 2.5 祝福语，sentence = sentence_1 + name + sentence_2 + selected_wish + sentence_3
#### sentence_1 = '祝'
#### sentence_2 = '新年快乐龙年大吉'
#### sentence_3 = '梁某人敬上二四年除夕'
#### wishes  # 存放随机祝福语的列表
## 3 运行
### flash_text_generation.py