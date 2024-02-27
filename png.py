import xlwings as xw
import time
import os, shutil
import zipfile
import configuration_file
from PIL import Image
from PIL import ImageGrab
import os

Image.MAX_IMAGE_PIXELS = 2300000000

start_time = time.time()

time.sleep(1)
shutil.copy2(configuration_file.Template_file, configuration_file.ls_file)
time.sleep(5)

# 模板所需发送文件截图，非截图区域隐藏，截图会从第一张开始往后开始
app = xw.App(visible=True, add_book=False)
wb = app.books.open(configuration_file.ls_file)
time.sleep(5)
i = 0
while i < configuration_file.Screets:  # Excel文件截图，所需张数9
    sheet = wb.sheets[i]
    false_list = []
    max_retry = 0
    while max_retry < 20:
        try:
            all = sheet.used_range
            time.sleep(2)
            all.api.CopyPicture()
            time.sleep(2)
            sheet.api.Paste()
            break
        except Exception:
            max_retry += 1
    i += 1
wb.save()
app.quit()

time.sleep(5)

configuration_file.del_file(configuration_file.ls1_folder)
time.sleep(1)


def compenent(excel_file_path, img_path):
    # 判断是否是文件和判断文件是否存在
    def isfile_exist(file_path):
        if not os.path.isfile(file_path):
            print("It's not a file or no such file exist ! %s" % file_path)
            return False
        else:
            return True

    # 修改指定目录下的文件类型名，将excel后缀名修改为.zip
    def change_file_name(file_path, new_type='.zip'):
        if not isfile_exist(file_path):
            return ''
        extend = os.path.splitext(file_path)[1]  # 获取文件拓展名
        if extend != '.xlsx' and extend != '.xlsm':
            print("It's not a excel file! %s" % file_path)
            return False
        file_name = os.path.basename(file_path)  # 获取文件名
        new_name = str(file_name.split('.')[0]) + new_type  # 新的文件名，命名为：xxx.zip
        dir_path = os.path.dirname(file_path)  # 获取文件所在目录
        new_path = os.path.join(dir_path, new_name)  # 新的文件路径
        if os.path.exists(new_path):
            os.remove(new_path)
        os.rename(file_path, new_path)  # 保存新文件，旧文件会替换掉
        return new_path  # 返回新的文件路径，压缩包

    # 解压文件
    def unzip_file(zipfile_path):
        if not isfile_exist(zipfile_path):
            return False
        if os.path.splitext(zipfile_path)[1] != '.zip':
            print("It's not a zip file! %s" % zipfile_path)
            return False
        # file_zip = zipfile.ZipFile(zipfile_path, 'r')
        file_zip = zipfile.ZipFile(zipfile_path)
        file_name = os.path.basename(zipfile_path)  # 获取文件名
        zipdir = os.path.join(os.path.dirname(zipfile_path), str(file_name.split('.')[0]))  # 获取文件所在目录
        for files in file_zip.namelist():
            file_zip.extract(files, os.path.join(zipfile_path, zipdir))  # 解压到指定文件目录
        file_zip.close()
        return True

    # 读取解压后的文件夹，打印图片路径
    def read_img(zipfile_path, img_path):
        if not isfile_exist(zipfile_path):
            return False
        dir_path = os.path.dirname(zipfile_path)  # 获取文件所在目录
        file_name = os.path.basename(zipfile_path)  # 获取文件名
        unzip_dir = os.path.join(dir_path, str(file_name.split('.')[0]))
        pic_dir = 'xl' + os.sep + 'media'  # excel变成压缩包后，再解压，图片在media目录
        pic_path = os.path.join(dir_path, str(file_name.split('.')[0]), pic_dir)
        file_list = os.listdir(pic_path)
        for file in file_list:
            filepath = os.path.join(pic_path, file)
            # print(filepath,img_path)
            shutil.move(filepath, img_path)
        os.unlink(zipfile_path)
        shutil.rmtree(unzip_dir)

    # 组合各个函数

    zip_file_path = change_file_name(excel_file_path)
    if not os.path.exists(img_path):
        os.mkdir(img_path)
    if zip_file_path != '':
        unzip_msg = unzip_file(zip_file_path)
        if unzip_msg:
            read_img(zip_file_path, img_path)


compenent(configuration_file.ls_file, configuration_file.ls1_folder)

time.sleep(5)

configuration_file.del_file(configuration_file.ls2_folder)
time.sleep(1)
addr = configuration_file.ls1_folder
os.chdir(configuration_file.ls1_folder)
faddr = os.path.abspath('..')
target = configuration_file.ls2_folder

names = os.listdir(addr)
for name in names:
    img_addr = addr + '\\' + name
    sim_name = name.split(".", 1)[0]
    tar_addr = target + '\\' + sim_name + '.png'
    img = Image.open(img_addr)
    img.save(tar_addr)

time.sleep(5)

if os.path.exists(configuration_file.ls_file):
    os.remove(configuration_file.ls_file)
# os.remove(configuration_file.ls_file)

end_time = time.time()
elapsed_time = round(end_time - start_time, 2)
print("处理播报数据转图片时间为：", elapsed_time, "秒")
