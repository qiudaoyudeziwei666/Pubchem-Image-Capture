import requests
from lxml import html
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PilImage
from io import BytesIO
import os
import time

# 获取用户输入的文件夹路径
#excel_file_path_input = input("请输入你的Excel输入文件路径：")
excel_file_path_input = r'C:\Users\fen\Desktop\LC (2)\正负离子石化厂，LCMS,GNPS.xlsx'
#excel_file_path_output = input("请输入你的Excel输出文件路径：")
excel_file_path_output = r'C:\Users\fen\Desktop\LC (2)\正负离子石化厂，LCMS,GNPS - 副本.xlsx'
#pig_path_output = input("请输入你保存图片的位置：")
pig_path_output = r'E:\pig'
# 使用 pandas 读取 Excel 文件
df = pd.read_excel(excel_file_path_input)

# 提取第一列的化学名称，假设第一列名称为 'Chemical Name'
chemical_names = df.iloc[:, 0].tolist()

# 打印化学名称列表
print("化学名称列表：", chemical_names)

# 打开输出Excel工作簿
workbook = load_workbook(excel_file_path_output)
sheet = workbook.active


# 图片插入列的起始行
start_row = 2

def fetch_image(url):
    response = requests.get(url)
    if response.status_code == 200:
        return BytesIO(response.content)
    else:
        return None

# 根据上面提供的化合物名称去网页中查找信息
counter = 0  # 计数器
for c_names in chemical_names:
    try:
        url = "https://pubchem.ncbi.nlm.nih.gov/#query=" + c_names
        # 设置 Chrome 选项
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # 可选：如果你不需要看到浏览器界面
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")

        # 初始化 WebDriver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)

        # 打开目标网页
        driver.get(url)

        # 等待页面完全加载
        time.sleep(5)  # 根据需要调整等待时间

        # 获取页面的HTML内容
        html_content = driver.page_source

        # 解析 HTML
        tree = html.fromstring(html_content)

        # 使用 XPath 找到包含目标 href 属性的 <a> 元素
        a_elements = tree.xpath("//a[contains(@href, 'pubchem.ncbi.nlm.nih.gov/compound')]")

        # 定义正则表达式模式以提取 href 中的数字
        number_regex = re.compile(r'/compound/(\d+)')
        # 遍历找到的元素并提取数字
        CID = None
        for element in a_elements:
            href_content = element.get("href")
            match = number_regex.search(href_content)
            if match != None:
                CID = match.group(1)
                print(f"Extracted number: {CID}")
                break
            else:
                print(f"第{start_row}行获取CID失败")

        # 关闭浏览器
        driver.quit()

        if CID:
            # 根据CID获取新的图片网址
            new_url = "https://pubchem.ncbi.nlm.nih.gov/image/imgsrv.fcgi?cid=" + CID + "&t=l"
            print(new_url)

            # 获取图片数据
            img_data = fetch_image(new_url)
            # 保存图片到输出目录
            if img_data:
                pil_img = PilImage.open(img_data)
                img_filename = pig_path_output + f"\\image_{CID}.png"
                pil_img.save(img_filename)
                print(str(CID) + "已经下载完成")

                # 加载图片并插入到Excel
                image_path = img_filename
                try:
                    img = Image(image_path)
                    print(f"图片加载成功: {image_path}")
                    cell = f"B{start_row}"
                    sheet.add_image(img, cell)
                    print(f"图片已经输入到第{start_row}行")
                except Exception as e:
                    print(f"图片加载失败: {e}")
                    sheet[f'B{start_row}'] = "找不到图片"
            else:
                sheet[f'B{start_row}'] = "找不到图片"
        else:
            sheet[f'B{start_row}'] = "找不到CID"


        start_row += 1
        counter += 1

        # 每处理10个化合物保存一次
        if counter == 10:
            workbook.save(excel_file_path_output)
            print(f"文件已经安全保存到{excel_file_path_output}")
            counter = 0  # 重置计数器

    except Exception as e:
        print(f"Error processing {c_names}: {e}")
        sheet[f'B{start_row}'] = "找不到图片"
        start_row += 1

# 最终保存修改后的工作簿
workbook.save(excel_file_path_output)
print(f"Final Excel file saved to {excel_file_path_output}")
