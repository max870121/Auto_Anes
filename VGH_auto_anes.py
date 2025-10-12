import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import os
import requests
from VGH_login import VGHLogin
from VGH_function import *
import pandas as pd
from bs4 import BeautifulSoup
import time
import random
import urllib.parse

import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT
from PIL import Image
from docx.oxml.ns import qn
import os
from datetime import datetime, timedelta
import pwinput
import platform

def clear_terminal():
    if platform.system() == "Windows":
        os.system('cls')
    else:
        os.system('clear')


def app_info():
    print("""
    此程式可以自動查詢病人資料，製作一份pdf 麻仿單
    請先打入入口網帳號密碼，之後輸入燈號或式直接按ENTER之後輸入病房
    ***注意 若病人太多可能會被資訊室鎖住該台電腦一陣子，可以重新開機後稍等一下
    作者的燈號為: 8375K，如果有任何問題或建議，歡迎聯絡!!!
    """)

def login_to_system():
    """處理登入邏輯"""
    clear_terminal()
    app_info()
    
    while True:
        # 要求使用者輸入帳號與密碼
        username = input("請輸入帳號：")
        password = pwinput.pwinput(prompt='密碼: ', mask='*')
        
        try:
            # 使用 VGHLogin 進行登入
            vgh = VGHLogin()
            if vgh.login(username, password):
                print("✅ 登入成功！")
                clear_terminal()
                return vgh
            else:
                clear_terminal()
                app_info()
                print(f"登入失敗:，請重新輸入帳號與密碼。\n")
        except Exception as e:
            clear_terminal()
            app_info()
            print(f"登入失敗: {e}，請重新輸入帳號與密碼。\n")

def add_text_to_pdf(input_pdf_path, output_pdf_path, text, x, y, page_num=0):
    # 讀取原 PDF
    with open(input_pdf_path, "rb") as input_pdf_file:
        reader = PyPDF2.PdfReader(input_pdf_file)
        writer = PyPDF2.PdfWriter()

        # 創建一個新的 PDF 文件來放置文字
        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        
        # 在指定位置 (x, y) 加上文字
        c.drawString(x, y, text)
        c.save()

        # 將文字加到原始 PDF 頁面上
        packet.seek(0)
        new_pdf = PyPDF2.PdfReader(packet)
        page = reader.pages[page_num]
        
        # 合併文字到指定頁面
        page.merge_page(new_pdf.pages[0])

        # 將原始頁面加回 PDF 編輯器
        writer.add_page(page)

        # 將其餘頁面也加入到新 PDF 文件中
        for i in range(1, len(reader.pages)):
            writer.add_page(reader.pages[i])

        # 輸出修改後的 PDF
        with open(output_pdf_path, "wb") as output_pdf_file:
            writer.write(output_pdf_file)

def add_text_with_wrap_to_pdf(input_pdf_path, output_pdf_path, text, x_text, y_text, width=200, fontSize=12,page_num=0):
    # 讀取原始 PDF
    with open(input_pdf_path, "rb") as input_pdf_file:
        reader = PyPDF2.PdfReader(input_pdf_file)
        writer = PyPDF2.PdfWriter()

        # 創建一個新的 PDF 文件來放置帶有換行的文字
        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        pdfmetrics.registerFont(TTFont('SimHei', 'mingliu.ttc')) 
        
        # 使用 reportlab 的樣式來格式化文本，這裡使用 getSampleStyleSheet() 提供的樣式
        styles = getSampleStyleSheet()
        # style = styles["Normal"]  # 使用 'Normal' 樣式
        c.setFont('SimHei', fontSize)
        custom_style = ParagraphStyle(name='CustomStyle', fontSize=fontSize, fontName='SimHei')
        
        # 生成 Paragraph 對象，處理換行
        paragraph = Paragraph(text, custom_style)
        
        # 設置文本的最大寬度，這裡設置為 400 像素
        paragraph.wrapOn(c, width, 600)  # 這裡設置文本框的寬度，可以根據需要調整
        
        # 渲染文本
        paragraph.drawOn(c, x_text, y_text)  # 設定文本顯示的左下角位置

        c.save()

        # 將帶有換行的文字加到原始 PDF 頁面上
        packet.seek(0)
        new_pdf = PyPDF2.PdfReader(packet)
        page = reader.pages[page_num]
        
        # 合併新內容到指定頁面
        page.merge_page(new_pdf.pages[0])

        # 將原始頁面加回 PDF 編輯器
        writer.add_page(page)

        # 將其餘頁面也加入到新 PDF 文件中
        for i in range(1, len(reader.pages)):
            writer.add_page(reader.pages[i])

        # 輸出修改後的 PDF
        with open(output_pdf_path, "wb") as output_pdf_file:
            writer.write(output_pdf_file)

def insert_table_into_pdf(input_pdf_path, output_pdf_path, df, x=50, y=500):
    """
    将 Pandas DataFrame 作为表格插入到现有的 PDF 文件中的指定位置。

    :param input_pdf_path: 要插入表格的现有 PDF 文件路径
    :param output_pdf_path: 输出合并后的 PDF 文件路径
    :param df: Pandas DataFrame 数据
    :param x: 表格在页面上的 x 坐标位置（默认 50）
    :param y: 表格在页面上的 y 坐标位置（默认 500）
    """
    # 将 pandas DataFrame 转换为 ReportLab 表格
    table_data = [df.columns.to_list()] + df.values.tolist()

    # 在内存中创建一个 BytesIO 对象来生成 PDF
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)

    # 创建一个 ReportLab 表格对象
    table = Table(table_data)

    # 添加表格样式
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 0),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    # 将表格添加到 PDF 文档
    doc.build([table])

    # 获取生成的 PDF
    buffer.seek(0)
    table_pdf = PyPDF2.PdfReader(buffer)

    # 打开现有的 PDF 文件
    existing_pdf = PyPDF2.PdfReader(input_pdf_path)

    # 创建一个 PdfWriter 对象
    pdf_writer = PyPDF2.PdfWriter()

    # 获取现有 PDF 中的第一页
    page = existing_pdf.pages[0]

    # 创建一个新的 PDF，并使用 reportlab.drawOn() 方法将表格绘制到指定位置
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    table_width, table_height = table.wrapOn(c, 400, 500)  # 获取表格尺寸 (可以根据需要调整)
    table.drawOn(c, x, y)  # 设定表格的插入位置 (x, y 为左下角坐标)

    # 完成并保存绘制的 PDF
    c.save()
    packet.seek(0)
    table_pdf_from_canvas = PyPDF2.PdfReader(packet)

    # 合并表格 PDF 和现有的 PDF
    page.merge_page(table_pdf_from_canvas.pages[0])

    # 将合并后的页面保存到新的 PDF 文件中
    pdf_writer.add_page(page)

    # 如果有更多页面，可以继续添加
    for i in range(1, len(existing_pdf.pages)):
        pdf_writer.add_page(existing_pdf.pages[i])

    # 保存最终的 PDF 文件
    with open(output_pdf_path, "wb") as output_file:
        pdf_writer.write(output_file)

vgh = login_to_system()
time.sleep(0.5)

page_content=vgh.get_page_after_login("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findEmr&histno=50687768")
soup = BeautifulSoup(page_content, 'html.parser')

print("登入成功，請輸入您要產生麻仿單的病歷號")
ID=input("病歷號:")

while not ID=="":
    try:
        OP=get_OP(vgh, ID)
        OP_Dx=OP["OP_Dx"]
        OP_Dx=OP_Dx.split()
        OP_Dx=" ".join(OP_Dx)
        if len(OP_Dx)>30:
            add_text_with_wrap_to_pdf("Lib/PreOP_Anes.pdf", "output.pdf", OP_Dx, 165, 697, 150, 10)
        else:
            add_text_with_wrap_to_pdf("Lib/PreOP_Anes.pdf", "output.pdf", OP_Dx, 165, 715, 150, 10)

        # #Surgery
        OP_name=OP["OP_name"]
        OP_name=OP_name.split()
        OP_name=" ".join(OP_name)
        if len(OP_name)>30:
            add_text_with_wrap_to_pdf("output.pdf", "output.pdf", OP_name, 425, 697, 150, 10)
        else:
            add_text_with_wrap_to_pdf("output.pdf", "output.pdf", OP_name, 425, 715, 150, 10)

        Anes=OP["Anes"]
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", Anes, 435, 610, 150, 10)
    except:
        pass

    
    # 病歷號
    try:
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", ID, 380, 764)

        admin_intro=get_admin_Intro(vgh,ID)

        # 床號
        Bed=admin_intro.at[0, "病房床號"]
        Bed=Bed.replace("－ ","-")
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", Bed, 445, 777)
        Bed=Bed.replace("-","_")

        # 生日年齡
        Age=admin_intro.at[0, "生　日　"].split("（")[1]
        Age=admin_intro.at[0, "生　日　"].split("（")[1]
        Age=Age.split("歲")[0]
        Birthday=admin_intro.at[0, "生　日　"].split("（")[0]
        Birthday_y=Birthday[0:4]
        Birthday_m=Birthday[4:6]
        Birthday_d=Birthday[6:]
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", Birthday_y, 390, 740)
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", Birthday_m, 445, 740)
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", Birthday_d, 495, 740)

        # 性別
        Sex=admin_intro.at[0, "性　別　"]

        pat_name=admin_intro.at[0, "姓　名　"].split("(")[0]
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", pat_name, 400, 750)

        

        # breakpoint()
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", Sex, 495, 752)
        # breakpoint()
        if "女" in Sex:
            Sex="F"
        else:
            Sex="M"
        
        Age=Age+"y/o "+Sex
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", Age, 55, 317)
    except:
        print("病歷號錯誤，請檢查")
        ID=input("病歷號:")
        continue

    try:
        BW_BL=get_BW_BL(vgh,ID, adminID="all")
        BW_BL="\\".join(list(BW_BL[["身高","體重","BMI"]].iloc[0]))
        add_text_with_wrap_to_pdf("output.pdf", "output.pdf", BW_BL, 55, 300)
    except:
        pass

    Hx="Hx"
    add_text_with_wrap_to_pdf("output.pdf", "output.pdf", Hx, 55, 280)


    try:
        SMAC=get_res_report(vgh,ID,resdtype="SMAC")
        SMAC=SMAC[["NA","K","BUN","CREA","GLU"]]
        SMAC=SMAC.tail(2)
        insert_table_into_pdf("output.pdf", "output.pdf", SMAC, x=300, y=300)
    except:
        pass
    
    try:
        CBC=get_res_report(vgh,ID,resdtype="CBC")
        CBC=CBC[["WBC","HGB","PLT", 'PT', 'APTT']]
        CBC=CBC.tail(2)
        insert_table_into_pdf("output.pdf", "output.pdf", CBC, x=300, y=230)
    except:
        pass

    ECG= "ECG:"
    add_text_with_wrap_to_pdf("output.pdf", "output.pdf", ECG, 330, 220)
    CXR= "CXR:"
    add_text_with_wrap_to_pdf("output.pdf", "output.pdf", CXR, 330, 200)
    name=Bed+"_"+ID+".pdf"
    try:
        os.rename('output.pdf', name)
        print("產生完成，儲存為:"+name)
    except:
        pass
    ID=input("病歷號:")