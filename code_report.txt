'''
Author:Zhang shaoniu
Email:shaoniu.zhang@cn.bosch.com
start date:2023/8/29 21:55
'''
import tkinter as tk
from tkinter import W
from tkinter import E
from tkinter import NW
from tkinter import N
from tkinter import Tk
from tkinter import ttk
from tkinter import END
from tkinter import Button
from tkinter import filedialog
import tkinter.messagebox
from tkinter import scrolledtext
from tkinter import simpledialog
from tkinter import Toplevel
from ComBoPicker import Combopicker
import time
import os
import shutil
from PIL import Image
from io import BytesIO
from selenium import webdriver  
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
import urllib.request
import requests
from requests.auth import HTTPBasicAuth
import threading
from openpyxl import load_workbook
import docx
from docx import Document
from docx.image.exceptions import UnrecognizedImageError
from docx.shared import Pt
from docx.shared import Cm  
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import RGBColor
from docx.parts.image import ImagePart
from docx.oxml.shape import CT_Picture
from docx.oxml import OxmlElement
from docx.enum.dml import MSO_THEME_COLOR_INDEX
from docx.text.paragraph import Paragraph
from lxml import etree
from docx.oxml.ns import qn

import win32con
import win32api
import win32com.client
import win32file
import win32clipboard
import pythoncom
import sys
import re
import zipfile
import subprocess

from datetime import datetime
from pathlib import Path
#import pandas as pd       #由于这几个包加载耗时比较长 所以改为在函数内import
# import matplotlib.pyplot as plt
# from matplotlib.ticker import AutoMinorLocator, AutoLocator
#import numpy as np
import json
from tkcalendar import Calendar

# if getattr(sys, 'frozen', False):
#     base = sys._MEIPASS
#     os.environ['TCL_LIBRARY'] = os.path.join(base, 'tcl')
#     os.environ['TK_LIBRARY'] = os.path.join(base, 'tk')


def show_current():
    ed=EdateApi()
    #a=ed.get_test_orders_by_filter(filter_name="zhang")
    a=ed.get_test_orders_by_filter("testinfo1","ZHS6CNG")
    #a=get_test_order_details(238783)
    #a=get_test_objects_by_entity_ids([238783])
    print(a)

def clear_input(a):
    a.clear()

def center_window(window):
    """居中显示窗口"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

def get_report_template(testid):
    pass

    #在进度栏中显示信息
def showinfo(st,string,colour):
    st.insert('end',string+'\n',colour)
    st.see(tkinter.END)        
#清除进度栏信息
def clear_status(st):
    st.delete(1.0, END)
#发送邮件
def send_email(email_title,recipient,email_content):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        print(f"Outlook 版本: {outlook.Version}")
        mail = outlook.CreateItem(0)
        mail.Subject = email_title
        mail.Body = email_content
        mail.To = recipient
        mail.Send()
        print("邮件发送完成！")
        return True
    except Exception as e:
        print(f"邮件发送失败: {e}")
        return False

#将 Excel 单元格位置（如 B3）转换为行号和列号
def excel_position_to_row_col(excel_position):
    """
    将 Excel 单元格位置（如 B3）转换为行号和列号。
    
    参数:
        excel_position (str): Excel 单元格位置，例如 'B3'
    
    返回:
        tuple: (行号, 列号)，例如 (3, 2)
    """
    # 分离字母部分和数字部分
    col_str = ''
    row_str = ''
    
    for char in excel_position:
        if char.isalpha():  # 如果是字母，则属于列号
            col_str += char
        elif char.isdigit():  # 如果是数字，则属于行号
            row_str += char
    # 计算列号
    col_num = 0
    for i, char in enumerate(col_str[::-1]):
        col_num += (ord(char.upper()) - ord('A') + 1) * (26 ** i)
    # 计算行号
    row_num = int(row_str)
    return row_num, col_num

#压缩文件夹里的所有文件
def zip_folder(source_folder, output_path=None):
    """
    将指定文件夹打包成zip文件
    
    Args:
        source_folder: 要打包的文件夹路径
        output_path: 输出zip文件的路径
            - 如果是目录，则在该目录下创建同名zip文件
            - 如果是文件路径，则使用该路径
            - 如果为None，则在源文件夹同级目录创建zip文件
    """
    source_path = Path(source_folder)
    
    if not source_path.exists() or not source_path.is_dir():
        print(f"错误: 文件夹 '{source_folder}' 不存在")
        return False
    
    # 设置输出文件路径
    if output_path is None:
        # 默认放在同级目录
        zip_name = f"{source_path.name}.zip"
        output_path = source_path.parent / zip_name
    else:
        output_path = Path(output_path)
        
        # 检查output_path是文件还是目录
        if output_path.exists() and output_path.is_dir():
            # 如果是目录，在该目录下创建同名zip文件
            zip_name = f"{source_path.name}.zip"
            output_path = output_path / zip_name
        else:
            # 确保zip文件有.zip扩展名
            if output_path.suffix.lower() != '.zip':
                output_path = output_path.parent / f"{output_path.name}.zip"
    
    try:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(source_path):
                for file in files:
                    file_path = Path(root) / file
                    # 在zip中保持相对路径结构
                    # 相对于源文件夹的路径
                    arcname = file_path.relative_to(source_path)
                    zipf.write(file_path, arcname)
        
        print(f"压缩完成: {output_path}")
        return output_path
        
    except Exception as e:
        print(f"压缩失败: {e}")
        return False

#寻找数据所在的根目录
def search_root_dir(path):
    while True:
        path, part = os.path.split(path)
        if check_pattern(part):
            return path+'\\'+part   
        if not path or path == os.sep:  # 到达根目录时终止
            break
    #遍历文件夹，返回所有文件名包含"string"且扩展名为.txt的文件路径列表
def search_file3(folder_path,string):
    """
    遍历文件夹，返回所有文件名包含"string"且扩展名为.txt的文件路径列表
    :folder_path: 要遍历的文件夹路径
    :return: 符合条件的文件路径列表
    """
    matched_files = []  # 存储匹配文件的数组
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            # 检查文件是否为.txt格式且文件名包含"string"（不区分大小写）
            if filename.endswith('.txt') and string in filename:
                full_path = os.path.join(root, filename)
                matched_files.append(full_path)
    return matched_files        

#检查字符串是否满足前3位数字+第4位下划线
def check_pattern(s):
    """检查字符串是否满足前3位数字+第4位下划线"""
    if len(s) < 4:  # 确保字符串长度足够
        return False
    # 同时验证前3字符是数字且第4字符是下划线
    return (s[0].isdigit() and 
            s[1].isdigit() and 
            s[2].isdigit() and 
            s[3] == '_')
    #获取某个文件的版本号
def get_file_version(file_path):
    if os.path.exists(file_path):  #如果文件存在
        try:
            # 尝试获取文件的版本信息
            info = win32api.GetFileVersionInfo(file_path, '\\')
            
            # 解析版本号
            ms = info['FileVersionMS'] # 高位版本号
            ls = info['FileVersionLS'] # 低位版本号
            
            # 版本号格式化
            version = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
            return version
        except win32api.error as e:
            if e.winerror == win32con.ERROR_BAD_EXE_FORMAT:
                return "The file is not an executable or does not have a version resource."
            else:
                return f"Error: {e}"
        except Exception as e:
            return f"An unexpected error occurred: {str(e)}"
    else:
        return None

#执行一个powershell命令        
def powershell_cmd(command):
    try:
        cmd_result=subprocess.run(["powershell","-Command",command],capture_output=True,text=True)
        result=cmd_result.stdout
    except Exception as e:
        result=f"Error:{str(e)}"
    return result

#获取某个盘符的UNC地址
def get_unc_path(drive_letter):
    result=os.popen('net use').read()
    lines=result.splitlines()
    for line in lines:
        if drive_letter+":" in line:
            parts=line.split(':')
            unc_path=parts[1].strip()
            return unc_path
    return None

    #保存共享的文件
def shared_pyxl_save(file_path,workbook,sheet_name):
    zin = zipfile.ZipFile(file_path, 'r')
    buffers = []
    for item in zin.infolist():
        #sheet2即Test Project Information对应位置
        if f"{sheet_name}.xml" not in item.filename:
            buffers.append((item, zin.read(item.filename)))
    zin.close()

    workbook.save(file_path)
    """ loop through again to find the sheet2.xmls and put it into buffer, else will show up error"""
    zin2 = zipfile.ZipFile(file_path, 'r')
    for item in zin2.infolist():
        if f"{sheet_name}.xml" in item.filename:
            buffers.append((item, zin2.read(item.filename)))
    zin2.close()
    
    #finally saves the file
    zout = zipfile.ZipFile(file_path, 'w')
    for item, buffer in buffers:
        zout.writestr(item, buffer)
    zout.close()
    workbook.close()


#从网络更新驱动
def update_browser_driver():
    showinfo(generate_report.report_status,'尝试从网络地址获取驱动.....','green')
    if download_driver()==True:
        showinfo(generate_report.report_status,'驱动下载完成','green')
        showinfo(generate_report.report_status,'正在解压文件','black')
        if(extract_file(driver_package_path,current_dir))==True:
            showinfo(generate_report.report_status,'解压成功','green')  
            shutil.copyfile(driver_path,local_driver_path)  #将下载解压的驱动复制到本地驱动保存地址
        else:
            showinfo(generate_report.report_status,'解压失败','green')  
            return False 
        showinfo(generate_report.report_status,'正在删除下载驱动压缩包...','black')
        if(delete_dir(current_dir+'/Driver_Notes'))==True:
            showinfo(generate_report.report_status,'删除Driver_Notes成功','green')
        else:
            showinfo(generate_report.report_status,'删除Driver_Notes失败','red')   
        if(delete_file(current_dir+'/edgedriver_win64.zip'))==True:
            showinfo(generate_report.report_status,'删除edgedriver_win64.zip成功','green')
        else:
            showinfo(generate_report.report_status,'删除edgedriver_win64.zip失败','red')   
        return True       
    else:
        showinfo(generate_report.report_status,'驱动下载失败失败,尝试从本地获取驱动...','red')
        try:
            shutil.copyfile(local_driver_path, driver_path)
            showinfo(generate_report.report_status,'驱动更新成功','green')
            return True
        except Exception as e:
            showinfo(generate_report.report_status,'驱动更新失败','red')
            return False

#下载驱动安装包
def download_driver():
    driver_entire_url=driver_basic_url+'/'+blowser_version+'/'+'edgedriver_win64.zip'
    showinfo(generate_report.report_status,f'驱动网络地址:{driver_entire_url}','black')
    try: 
        response=requests.get(driver_entire_url)
        if response.status_code==200:
            with open('edgedriver_win64.zip','wb') as f:
                f.write(response.content) 
                return True
    except Exception as e:
        print(str(e))
        return False    

#解压一个压缩包
def extract_file(tar_file,dest_dir):
    try:
        with zipfile.ZipFile(tar_file,'r') as zip_ref:
            zip_ref.extractall(dest_dir)
        return True
    except Exception as e:
        return False

#删除一个文件
def delete_file(tar_file):
    if os.path.exists(tar_file):
        try:
            os.remove(tar_file)
            return True
        except Exception as e:
            return False
    else:
        print('file not exists')
        return False        

#删除一个目录和目录下面的所有文件
def delete_dir(tar_dir):
    if os.path.exists(tar_dir):
        try:
            shutil.rmtree(tar_dir)
            return True
        except Exception as e:
            return False
    else:
        print('dir not exists')
        return False                

#更新设备计量信息
def update_calibration_info():
    global equipment_file_version
    global equipment_list
    print('formal equipment_file_version:'+equipment_file_version)
    file_list=os.listdir(equipment_file_path)
    if len(file_list)!=0:
        for file in file_list:
            if file[0]!='~':  #transfer盘中可能会存在隐藏文件 ，必须使用判断文件名首个字符为~来排除隐藏文件,无法用file_is_hidden方法
                current_equipment_file_version=file[9:].split('.')[0]
                print('current_equipment_file_version:'+current_equipment_file_version)
                if current_equipment_file_version!=equipment_file_version:
                    print('need to update the calibration information')
                    excel_wb=load_workbook(equipment_file_path+'/'+file,read_only=True)
                    excel_ws=excel_wb.active              #选择活动的工作表
                    for i,equipment in enumerate(equipment_list):
                        equipment_name=equipment.split('/',1)[0]                             #设备名
                        calibration_id=equipment.split('/',1)[1].split('(')[0]               #计量编号
                        calibration_date=equipment.split('/',1)[1].split('(')[1].split(')')[0]  #计量有效期
                        print('before update:'+equipment)
                        for row in excel_ws.iter_rows():
                            if row[1].value==calibration_id:
                                calibration_date=row[14].value[0:10]
                                equipment_list[i]=equipment_name+'/'+calibration_id+'('+calibration_date+')'
                                print('after update:'+equipment_list[i])   
                    print('update the equipment_file_version from '+equipment_file_version+'------>'+current_equipment_file_version)               
                    equipment_file_version=current_equipment_file_version  
                    excel_wb.close()
                    return True            
                else:
                    print('no need to update the calibration information')   
                    return None        
    else:
        print('equipment_file not exist')           
        return False                               

#中断程序执行
def interrupt_operation():
    global max_progress_value
    max_progress_value=0
    sys.exit(1)

#给一段文件添加超链接    
def add_hyperlink(paragraph, text, url):
# 获取paragraph的part对象，并且创建一个新的关系
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
# 创建 w:hyperlink 标签并且添加所需的值
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )
# 创建一个 w:r 元素 和一个 w:rPr 元素
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')
# 把所有的 xml 元素合并到一起，并且添加显示的文本给 w:r 元素
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
# 创建一个新的 Run 对象并且添加新的超链接进去
    r = paragraph.add_run()
    r._r.append(hyperlink)
# 设置超链接的样式，样式颜色为超链接，并且文本有下划线
    r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    r.font.underline = True
    return hyperlink


def select_config():
    selected_file = tkinter.filedialog.askopenfilename(title='选择配置文件', filetypes=[('所有文件', '.*')])
    if not os.path.basename(selected_file).startswith('config'):
        showinfo(generate_report.report_status,'所选配置文件不正确',"red")
    else:
        load_config(selected_file)

def load_config(config_path):
    """从JSON文件加载配置到全局变量"""
    global download_dir, plan_file, test_dir, raw_data_dir
    global approval_list, conclusion_list, equipment_list, folder_list, column_list,recipients_list
    global first_approval_current_value, second_approval_current_value, cobo_team_name_current_value,cobo_test_type_current_value
    global sheet_name_current_value, begin_row_current_value, testinfo_download_current_value
    global testinfo_plan_current_value, object_plan_current_value, addition_info_current_value
    global window_width, window_height
    global update_testinfo_current_value, create_folder_current_value, download_spec_current_value,with_foldername_current_value,with_foldername1_current_value
    global search_test_criteria_current_value, search_object_criteria_current_value,search_test_criteria1_current_value
    global compress_picture_current_value, time_sort_current_value, curve_generate_current_value,need_report_current_value
    global hide_page_current_value, cobo_register_type_current_value
    global equipment_file_version, data_character_current_value, sampling_frequency_current_value
    global x_unit_current_value, coordinate_typle_current_value, hyperlink_column_current_value
    global overwrite_current_value
    
    try:
        if not os.path.exists(config_path):
            print(f"配置文件不存在: {config_path}")
            return False
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        # 地址
        #download_dir = config_data.get("download_dir", "")                                                #下载地址
        plan_file = config_data.get("plan_file", "")                                                      #测试计划表地址
        test_dir = config_data.get("test_dir", "")                                                        #数据文件夹地址
        raw_data_dir = config_data.get("raw_data_dir", "")                                                #原始数据文件夹地址
        # 报告生成选项
        report_options = config_data.get("report_options", {})
        hide_page_current_value = report_options.get("hide_page_current_value", "")               #是否隐藏抓取页面
        compress_picture_current_value = report_options.get("compress_picture_current_value", "")                           #是否压缩图片
        time_sort_current_value = report_options.get("time_sort_current_value", "")                                 #是否时间排序
        curve_generate_current_value = report_options.get("curve_generate_current_value", "")                            #是否生成曲线
        need_report_current_value=report_options.get("need_report_current_value", "")                      #是否需要报告
        #设备
        equipment_list = config_data.get("equipment_list", [])                                            #设备信息
        equipment_file_version = config_data.get("equipment_file_version", "")                            #设备清单版本
        #结论
        conclusion_list = config_data.get("conclusion_list", [])                                          #结论信息
        #审批人
        approvals_data = config_data.get("approval_data", {})                                             #审批人信息
        approval_list = approvals_data.get("approval_list", [])
        first_approval_current_value = approvals_data.get("first_approval_current_value", "")                           #审批人1默认值
        second_approval_current_value = approvals_data.get("second_approval_current_value", "")                           #审批人2默认值
        #收件人 
        recipients_list = config_data.get("recipients_list", [])                                          #收件人信息
         # 窗口大小
        window_size = config_data.get("window_size", {})                      #软件界面尺寸
        window_width = window_size.get("width", "")
        window_height = window_size.get("height", "")
        # 更新选项
        update_options = config_data.get("update_options", {})
        cobo_test_type_current_value = update_options.get("cobo_test_type_current_value", "")                 #试验类型
        if cobo_test_type_current_value.isdigit():
            cobo_test_type_current_value = int(cobo_test_type_current_value)
        else:
            cobo_test_type_current_value = cobo_test_type_current_value
        cobo_register_type_current_value = update_options.get("cobo_register_type_current_value", "")               #登记类型
        if cobo_register_type_current_value.isdigit():
            cobo_register_type_current_value = int(cobo_register_type_current_value)
        else:
            cobo_register_type_current_value = cobo_register_type_current_value
        update_testinfo_current_value = update_options.get("update_testinfo_current_value", 1)                              #是否更新测试信息表
        create_folder_current_value = update_options.get("create_folder_current_value", 1)                              #是否创建测试文件夹
        download_spec_current_value = update_options.get("download_spec_current_value", 1)                              #是否下载测试规范
        with_foldername_current_value = update_options.get("with_foldername_current_value", 1)                              #是否项目文件夹附带名称
        with_foldername1_current_value = update_options.get("with_foldername1_current_value", 1)                              #是否测试文件夹附带名称
        search_test_criteria_current_value = update_options.get("search_test_criteria_current_value", "")                           #测试信息搜索条件名1
        search_object_criteria_current_value = update_options.get("search_object_criteria_current_value", "")                           #样品信息搜索条件名
        search_test_criteria1_current_value = update_options.get("search_test_criteria1_current_value", "")                           #测试信息搜索条件名2
        sheet_name_current_value = update_options.get("sheet_name_current_value", "")                           #计划表分页
        begin_row_current_value = update_options.get("begin_row_current_value", "")                           #计划表起始行
        hyperlink_column_current_value = update_options.get("hyperlink_column_current_value", "")                         #超链接列号
        testinfo_download_current_value = update_options.get("testinfo_download_current_value", "")                           #测试信息(下载表)
        testinfo_plan_current_value = update_options.get("testinfo_plan_current_value", "")                           #测试信息(计划表)
        object_plan_current_value = update_options.get("object_plan_current_value", "")                           #样品信息(计划表)
        addition_info_current_value = update_options.get("addition_info_current_value", "")                           #附加信息(计划表)
        folder_list = update_options.get("folder_list", [])
        
        # 曲线生成选项
        curve_options = config_data.get("curve_options", {})
        data_character_current_value = curve_options.get("data_character_current_value", "")                                     #数据特征
        sampling_frequency_current_value = curve_options.get("sampling_frequency_current_value", "")                                    #采样频率
        column_list = curve_options.get("column_list", [])                                                #数据列特征
        x_unit_current_value = curve_options.get("x_unit_current_value", "")                                      #X轴单位
        coordinate_typle_current_value = curve_options.get("coordinate_typle_current_value", "")                                       #坐标轴类型
        
        #是否覆盖原数据
        overwrite_current_value = config_data.get("overwrite_current_value", "")                  #是否覆盖原数据
        cobo_team_name_current_value=config_data.get("cobo_team_name_current_value", "")                          #小组名称
        print(f"配置已从 {config_path} 加载")
        return True
        
    except json.JSONDecodeError as e:
        print(f"JSON解析错误: {e}")
    except Exception as e:
        print(f"加载配置失败: {e}")
        return False

def save_config(config_path):
    """保存配置到JSON文件"""
    try:
        # 导入需要的模块（根据您的实际代码调整）
        # from your_module import generate_report, update_test_plan, test_curve, data_copy
        # from your_module import get_window_size, StringVarlist5, StringVarlist6
        # 构建配置字典
        fl=[]
        cl=[]
        for s in StringVarlist5:
            fl.append(s.get())
        for s in StringVarlist6:
            cl.append(s.get())    
        config_data = {
            "download_dir": download_dir,
            "plan_file": plan_file,
            "test_dir": test_dir,
            "raw_data_dir": raw_data_dir,
            "report_options": {
                "hide_page_current_value": generate_report.hide_page.get(),
                "compress_picture_current_value": generate_report.compress_picture.get(),
                "time_sort_current_value": generate_report.time_sort.get(),
                "curve_generate_current_value": generate_report.curve_generate.get(),
                "need_report_current_value":generate_report.need_report.get()
            },
            "equipment_list": equipment_list,
            "equipment_file_version": equipment_file_version,
            "conclusion_list": conclusion_list,
            "approval_data": {
                "approval_list": approval_list,
                "first_approval_current_value": generate_report.first_approval.current(),
                "second_approval_current_value": generate_report.second_approval.current()
            },
            "recipients_list": recipients_list,
            "window_size": {
                "width": str(get_window_size()[0]),
                "height": str(get_window_size()[1])
            },
            "update_options": {
                "cobo_test_type_current_value": str(update_test_plan.cobo_test_type.current()),
                "cobo_register_type_current_value": str(update_test_plan.cobo_register_type.current()),
                "update_testinfo_current_value": update_test_plan.update_testinfo.get(),
                "create_folder_current_value": update_test_plan.create_folder.get(),
                "download_spec_current_value": update_test_plan.if_download_spec.get(),
                "with_foldername_current_value": update_test_plan.with_foldername.get(),
                "with_foldername1_current_value": update_test_plan.with_foldername1.get(),
                "search_test_criteria_current_value": update_test_plan.search_test_criteria.get(),
                "search_object_criteria_current_value": update_test_plan.search_object_criteria.get(),
                "search_test_criteria1_current_value": update_test_plan.search_test_criteria1.get(),
                "sheet_name_current_value": update_test_plan.sheet_name.get(),
                "begin_row_current_value": update_test_plan.begin_row.get(),
                "hyperlink_column_current_value": update_test_plan.hyperlink_column.get(),
                "testinfo_download_current_value": update_test_plan.testinfo_download.get(),
                "testinfo_plan_current_value": update_test_plan.testinfo_plan.get(),
                "object_plan_current_value": update_test_plan.object_plan.get(),
                "addition_info_current_value": update_test_plan.addition_info.get(),
                "folder_list": fl
            },
            "curve_options": {
                "data_character_current_value": test_cuve.data_character.get(),
                "sampling_frequency_current_value": test_cuve.sampling_frequency.get(),
                "column_list": cl,
                "x_unit_current_value": test_cuve.x_unit.get(),
                "coordinate_typle_current_value": test_cuve.coordinate_typle.get()
            },
            "overwrite_current_value": data_copy.overwrite.get(),
            "cobo_team_name_current_value":data_copy.cobo_team_name.current(),              
            "version": "1.0",
            "metadata": {
                "last_saved": "2024-01-01T00:00:00",  # 您可以使用datetime.now().isoformat()
                "application": "Your Application Name"
            }
        }
        
        # 确保目录存在
        config_dir = os.path.dirname(config_path)
        if config_dir and not os.path.exists(config_dir):
            os.makedirs(config_dir)
        
        # 保存到文件
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2, ensure_ascii=False)
        
        print(f"配置已保存到 {config_path}")
        return True
        
    except Exception as e:
        print(f"保存配置失败: {e}")
        return False

#初始化浏览器
def initial_browser(website_url):
    global driver
    global edge_options   
    if edge_options is None:
        edge_options = Options()
        a,b=get_screen_size()
        edge_options.add_argument(f"--window-size={a},{b}")  #窗口大小
        if generate_report.hide_page.get()==1:
            edge_options.add_argument("--window-position=0,-2000")   #窗口位置         
    if driver is None:
        if driver_version != blowser_version or driver_version==None:   
            showinfo(generate_report.report_status,'浏览器版本与驱动版本不一致或者驱动不存在，正在更新驱动中...',"red")
            update_browser_driver()
            time.sleep(3)
        driver=webdriver.Edge(options=edge_options)
        showinfo(generate_report.report_status,'打开浏览器中...',"black")
        driver.refresh()    
        driver.get(website_url)    
    else:           
        showinfo(generate_report.report_status,'打开浏览器中...',"black")
        driver.refresh()    
        driver.get(website_url)    
        
#将一组testid按照同一序列和不同序列分开    
def classify_testid(testidlist):
    idlist1 = []  # 存储不同测试序列的testid
    idlist2 = testidlist.copy()  # 存储相同序列的testid
    for i in range(len(testidlist)):
        e = 1
        for j in range(len(testidlist)):
            if i == j:
                continue
            if testidlist[i].split('.')[0] == testidlist[j].split('.')[0]:
                e = 0
                break
        if e == 1:
            idlist1.append(testidlist[i])
    for k in range(len(idlist1)):
        idlist2.remove(idlist1[k])
    return idlist1,idlist2    

#在指定excel表中查找指定key_index值为key_value的数据列并返回该数据列指定的index值
def get_value(excel_sheet,key_index,key_value,index):
    max_row=excel_sheet.max_row
    for i in range(max_row):
        if key_value==excel_sheet.cell(i+1,key_index).value:
            return excel_sheet.cell(i+1,index).value
    return None    

#获取equipment_list
def get_equipment_list(excel_path,key_index,key_value):
    results=[]
    try:
        excel_wb=load_workbook(excel_path,read_only=True)
        excel_ws=excel_wb.active              #选择活动的工作表
        #excel_ws=excel_wb['量检具导出列表']    #选择sheet名字的sheet
        for row in excel_ws.iter_rows():       #用这种方式遍历excel 表格必须把前面合并单元格的行先删除 否则运行报错 原因未知
            if row[key_index-1].value==key_value:
                result=row[0].value+'/'+row[1].value+'('+row[14].value+')'
                results.append(result)    
        excel_wb.close()        
        return results        
    except Exception as e:
        print('error in reading the excel')
        print(str(e))
        excel_wb.close()
        return results

#获取指定excel sheet当前的最大行号
def get_max_row(excel_sheet):
    for i in range(excel_sheet.max_row, -1, -1): 
        if excel_sheet.cell(i, 2).value is not None:
            return i
        

#判断页面元素是否存在    
def isElementExist(element):
    global driver
    flag=True
    try:
        driver.find_element(By.XPATH,element)
        return flag
    except:
        flag=False
        return flag  
    
#判断某个页面元素下是否含有特定text
def is_text_in_element(element,string):
    text=element.text
    if string in text:
        return True
    else:
        return False
    
#等待页面某个元素消失
def wait_element_invisibility(element,element_name,timeout):
    try:
        showinfo(generate_report.report_status,f'等待{element_name}图标消失...',"black")
        WebDriverWait(driver,timeout, 0.5).until(EC.invisibility_of_element_located(element))
        showinfo(generate_report.report_status,f'{element_name}图标已消失...',"black")
    except Exception as e:
        showinfo(generate_report.report_status,f'等待{element_name}元素消失时超时',"red")

#等待页面某个元素出现
def wait_element_visibility(element,element_name,timeout):
    try:
        showinfo(generate_report.report_status,f'等待{element_name}图标消失...',"black")
        WebDriverWait(driver,timeout, 0.5).until(EC.presence_of_element_located(element))
        showinfo(generate_report.report_status,f'{element_name}元素已出现...',"black")
    except Exception as e:
        showinfo(generate_report.report_status,f'等待{element_name}元素出现时超时',"red")

#等待页面某个元素含有特定text        
def wait_element_text(element,element_name,text,timeout):
    try:
        showinfo(generate_report.report_status,f'等待显示{element_name}...',"black")
        WebDriverWait(driver, timeout, 0.5).until(EC.text_to_be_present_in_element(element,text))
        showinfo(generate_report.report_status,f'已显示{element_name}...',"black")
    except Exception as e:
        showinfo(generate_report.report_status,f'等待{element_name}显示时超时',"red")

#等待页面某个元素可以被点击    
def wait_element_clickable(element,element_name,timeout):
    try:
        showinfo(generate_report.report_status,f'等待{element_name}图标完全加载并可被点击...',"black")
        WebDriverWait(driver, timeout, 0.5).until(EC.element_to_be_clickable(element))
        
    except Exception as e:
        showinfo(generate_report.report_status,f'等待{element_name}元素可被点击时超时',"red")
        
#点击窗口右上角退出程序同时将浏览器窗口关闭    
def close_window():  
    global driver
    window.destroy()
    if driver is not None:
        driver.quit()

#获取tkinter窗口大小        
def get_window_size():
    #获取缩放后tkinter窗口大小
    width=window.winfo_width()
    height=window.winfo_height()
    return width,height

#获取当前电脑显示屏尺寸
def get_screen_size():   
    w=window.winfo_screenwidth()
    h=window.winfo_screenheight()
    return w,h

#获取使得窗口居中的坐标轴
def get_middle_position(par_window,sub_window):
    x = par_window.winfo_x() + (par_window.winfo_width() // 2) - (sub_window.winfo_width() // 2)
    y = par_window.winfo_y() + (par_window.winfo_height() // 2) - (sub_window.winfo_height() // 2)    
    return x,y

# #刷新进度条        
# def update_progress():
#     global current_progress
#     global max_progress_value
#     if current_progress<max_progress_value:
#         current_progress+=0.05
#         progressbar["value"]=current_progress
#         status_frame.after(100,update_progress)


#判断文件是否被打开
def is_file_locked(filepath):
    locked = None
    try:
        handle = win32file.CreateFile(filepath, win32con.GENERIC_READ, win32con.FILE_SHARE_READ, None, win32con.OPEN_EXISTING, win32con.FILE_ATTRIBUTE_NORMAL, None)
        handle.close()
    except:
        locked = True
    else:
        locked = False
    return locked

#通过文件后缀判断是否为图片
def is_file_image(filepath):
    image_extensions=['.png','.jpg','.jpeg','.gif','.bmp','.tiff']
    ext=os.path.splitext(filepath)[1].lower()
    return ext in image_extensions

#判断是否隐藏文件或系统文件
def file_is_hidden(file):  
    attribute = win32api.GetFileAttributes(file)
    return attribute & (win32con.FILE_ATTRIBUTE_HIDDEN | win32con.FILE_ATTRIBUTE_SYSTEM)

#判断是否是excel文件
def is_excel_file(filename): 
    ext=os.path.splitext(filename)[1]
    if ext in['.xlsx','.xls']:
        return True
    else:
        return False


#增加设备
def add_equipment(window):
    add_equipment_window=Toplevel(window)
    w,h=get_screen_size()
    x=w//2-150
    y=h//2-50
    add_equipment_window.geometry(f"300x100+{x}+{y}")
    add_equipment_window.title('增加设备信息')
    label1=tkinter.Label(add_equipment_window,text="设备名称:")
    label1.grid(row=1,column=1)
    entry1=tkinter.Entry(add_equipment_window)
    entry1.grid(row=1,column=2)    

    label2=tkinter.Label(add_equipment_window,text="计量编号:")
    label2.grid(row=2,column=1)
    entry2=tkinter.Entry(add_equipment_window)
    entry2.grid(row=2,column=2) 

    label3=tkinter.Label(add_equipment_window,text="校准有效期:")
    label3.grid(row=3,column=1)
    entry3=tkinter.Entry(add_equipment_window)
    entry3.grid(row=3,column=2) 

    button1=tkinter.Button(add_equipment_window,text="确定",command=lambda:submit_add_equipment(entry1,entry2,entry3))
    button1.grid(row=4,column=1)
    #add_equipment_window.mainloop()

#提交增加的设备
def submit_add_equipment(a,b,c):
    if a.get()!='' and b.get()!='' and c.get()!='':
        equipment=a.get()+'/'+b.get()+'('+c.get()+')'
        equipment_list.append(equipment)
        a.delete(0,END)
        b.delete(0,END)
        c.delete(0,END)
    else:
        tkinter.messagebox.showinfo('提示', '输入空值无效')


def select_delete(window,window_title,option_list):
    select_delete_window = Toplevel(window)  # 用Toplevel方法创建子窗口
    w, h = get_screen_size()
    x = w // 2 - 210
    y = h // 2 - 225
    select_delete_window.geometry(f"420x450+{x}+{y}")
    select_delete_window.title(window_title)
    
    # 允许窗口调整大小
    select_delete_window.resizable(True, True)
    
    # 创建主框架
    main_frame = tkinter.Frame(select_delete_window)
    main_frame.pack(fill=tkinter.BOTH, expand=True, padx=5, pady=5)
    
    # 创建画布和滚动条
    canvas = tkinter.Canvas(main_frame, width=400, height=400)
    
    # 创建垂直和水平滚动条
    v_scrollbar = tkinter.Scrollbar(main_frame, orient=tkinter.VERTICAL, command=canvas.yview)
    h_scrollbar = tkinter.Scrollbar(main_frame, orient=tkinter.HORIZONTAL, command=canvas.xview)
    
    # 创建可滚动的内部框架
    inner_frame = tkinter.Frame(canvas)
    
    # 配置画布
    canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
    canvas.create_window((0, 0), window=inner_frame, anchor="nw")
    
    # 布局
    canvas.grid(row=0, column=0, sticky="nsew")
    v_scrollbar.grid(row=0, column=1, sticky="ns")
    h_scrollbar.grid(row=1, column=0, sticky="ew")
    
    # 配置网格权重
    main_frame.grid_rowconfigure(0, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)
    
    j = 0
    varlist = []
    for i, option in enumerate(option_list):
        var = tkinter.BooleanVar()
        checkbutton = tkinter.Checkbutton(inner_frame, text=option, variable=var, 
                                         onvalue=True, offvalue=False)
        checkbutton.grid(row=i, column=0, sticky=W, padx=10, pady=5)
        varlist.append(var)
        j = i + 1
    
    # 创建按钮框架（放在画布外部）
    button_frame = tkinter.Frame(main_frame)
    button_frame.grid(row=2, column=0, columnspan=2, pady=10)
    
    button3 = tkinter.Button(button_frame, text="全选", 
                            command=lambda: select_all(varlist))
    button3.pack(side=tkinter.LEFT, padx=10)

    button1 = tkinter.Button(button_frame, text="确定删除", 
                            command=lambda: submit_delete(option_list,varlist, select_delete_window))
    button1.pack(side=tkinter.LEFT, padx=10)
    
    button2 = tkinter.Button(button_frame, text="取消", 
                            command=select_delete_window.destroy)
    button2.pack(side=tkinter.LEFT, padx=10)
   
    
    # 更新内部框架大小并配置滚动区域
    inner_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))
    
    # 绑定鼠标滚轮事件
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def _on_shift_mousewheel(event):
        canvas.xview_scroll(int(-1*(event.delta/120)), "units")
    
    # 绑定鼠标滚轮事件
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)
    
    # 窗口关闭时解绑事件
    def on_close():
        canvas.unbind_all("<MouseWheel>")
        canvas.unbind_all("<Shift-MouseWheel>")
        select_delete_window.destroy()
    
    select_delete_window.protocol("WM_DELETE_WINDOW", on_close)    

#提交删除的设备
def submit_delete(option_list,varlist,select_delete_window):
    global generate_report
    for index in range(len(varlist)-1,-1,-1):
        if varlist[index].get()==True:
            option_list.pop(index)
    select_delete_window.destroy()
    generate_report.first_approval['values']=approval_list
    generate_report.second_approval['values']=approval_list

#全选
def select_all(varlist):
    for index in range(len(varlist)-1,-1,-1):
        if varlist[index].get()==False:
            varlist[index].set(True)

#增加结论
def add_conclusion(window,string):
    add_conclusion_window=Toplevel(window)
    w,h=get_screen_size()
    x=w//2-150
    y=h//2-50
    add_conclusion_window.geometry(f"300x100+{x}+{y}")
    add_conclusion_window.title('增加结论描述')
    label1=tkinter.Label(add_conclusion_window,text="结论描述:")
    label1.grid(row=1,column=1)
    text1=tkinter.Text(add_conclusion_window,height=4,width=28)
    text1.grid(row=1,column=2)    

    button1=tkinter.Button(add_conclusion_window,text="确定",command=lambda:submit_add_conclusion(text1,string))
    button1.grid(row=4,column=1)

#提交增加的结论
def submit_add_conclusion(a,string):
    global conclusion_list
    if a.get("1.0","end-1c")!='':
        conclusion_list.append(a.get("1.0","end-1c"))
        a.delete('1.0',END)
    else:
        tkinter.messagebox.showinfo('提示', '输入空值无效')
    

#增加审批人
def add_approval(window):
    add_approval_window=Toplevel(window)
    w,h=get_screen_size()
    x=w//2-150
    y=h//2-50
    add_approval_window.geometry(f"300x100+{x}+{y}")
    add_approval_window.title('增加审批人')
    label1=tkinter.Label(add_approval_window,text="审批人姓名:")
    label1.grid(row=1,column=1)
    entry1=tkinter.Entry(add_approval_window)
    entry1.grid(row=1,column=2)    

    label2=tkinter.Label(add_approval_window,text="审批人部门:")
    label2.grid(row=2,column=1)
    entry2=tkinter.Entry(add_approval_window)
    entry2.grid(row=2,column=2) 

    button1=tkinter.Button(add_approval_window,text="确定提交",command=lambda:submit_add_approval(entry1,entry2))
    button1.grid(row=4,column=1)

#提交增加的审批人
def submit_add_approval(a,b):
    global approval_list
    global generate_report
    if a.get()!='' and b.get()!='':
        approval=a.get()+'('+b.get()+')'
        approval_list.append(approval)
        a.delete(0,END)
        b.delete(0,END)
        generate_report.first_approval['values'] = approval_list
        generate_report.second_approval['values'] = approval_list
    else:
        tkinter.messagebox.showinfo('提示', '输入空值无效')

#增加收件人
def add_recipient(window):
    add_recipient_window=Toplevel(window)
    w,h=get_screen_size()
    x=w//2-150
    y=h//2-50
    add_recipient_window.geometry(f"300x100+{x}+{y}")
    add_recipient_window.title('增加收件人')
    label1=tkinter.Label(add_recipient_window,text="收件人姓名:")
    label1.grid(row=1,column=1)
    entry1=tkinter.Entry(add_recipient_window)
    entry1.grid(row=1,column=2)    
    button1=tkinter.Button(add_recipient_window,text="确定提交",command=lambda:submit_add_recipient(entry1))
    button1.grid(row=3,column=1)

#提交增加的收件人
def submit_add_recipient(a):
    global recipients_list
    global generate_report
    if a.get()!='':
        recipient=a.get()
        recipients_list.append(recipient)
        a.delete(0,END)
        #generate_report.first_approval['values'] = approval_list
    else:
        tkinter.messagebox.showinfo('提示', '输入空值无效')


def view_address():
    address_list=[]
    address_list.append(f"测试计划表地址:{plan_file}")
    address_list.append(f"测试数据文件夹地址:{test_dir}")
    address_list.append(f"原始数据文件夹地址:{raw_data_dir}")
    view_data(address_list,"各种地址信息")
    
def view_data(data_list, string):
    w, h = get_screen_size()
    x = w // 2 - 400
    y = h // 2 - 100
    
    # 创建新窗口
    list_window = tkinter.Toplevel(window)
    list_window.title(f'{string}')
    list_window.geometry(f"500x400+{x}+{y}")
    list_window.resizable(True, True)
    
    # 创建ScrolledText（自带垂直滚动条）
    st = scrolledtext.ScrolledText(list_window, wrap=tkinter.NONE)
    st.pack(fill=tkinter.BOTH, expand=True, padx=5, pady=5)
    
    # 添加水平滚动条
    h_scrollbar = tkinter.Scrollbar(st, orient=tkinter.HORIZONTAL, command=st.xview)
    h_scrollbar.pack(side=tkinter.BOTTOM, fill=tkinter.X)
    st.configure(xscrollcommand=h_scrollbar.set)
    
    # 插入数据
    for value in data_list:
        st.insert(tkinter.END, f"{value}\n")
    
    # 设置为只读
    st.configure(state=tkinter.DISABLED)
    
    return list_window


#选择测试计划表文件
def select_plan_file():  
    global plan_file
    path = tkinter.filedialog.askopenfilename(title='选择一个文件', filetypes=[('所有文件', '.*')]).replace('/','\\')
    if path!='':
        plan_file=path
        

#选择测试文件夹根目录
def select_test_dir():  
    global test_dir
    path = tkinter.filedialog.askdirectory().replace('/','\\')
    if path!='':
        test_dir=path
        data_copy.dst_folder.set(path)

#选择原始数据目录
def select_raw_data_dir():
    global raw_data_dir
    path = tkinter.filedialog.askdirectory().replace('/','\\')
    if path!='':
        raw_data_dir=path
        data_copy.src_folder.set(path)



#移动某个文件
def move_file(src_file, tar_dir,index):  
    if not os.path.exists(src_file):
        showinfo(generate_report.report_status,f'{src_file}源文件不存在',"red")
    elif not os.path.exists(tar_dir):
        showinfo(generate_report.report_status,f'{tar_dir}目标地址不存在',"red")
    else:  
        # 获取文件名
        filename = os.path.basename(src_file)
        target_path = os.path.join(tar_dir, filename)
        
        # 如果目标文件已存在
        if os.path.exists(target_path):
            if index==1:            #覆盖目标文件
                try:
                    os.remove(target_path)
                except PermissionError:
                    # 如果是只读文件，修改权限后再删除
                    os.chmod(target_path, 0o666) # 修改为可写
                    os.remove(target_path)
            else:                   #将原文件改名加时间后缀以实现不覆盖目标文件
                nowtime = time.strftime("%y-%m-%d-%M-%S", time.localtime())
                file_suffix = os.path.splitext(src_file)[-1]  # 获取文件后缀
                file_text = os.path.splitext(src_file)[0]
                new_src_file = file_text + '_' + nowtime + file_suffix  # 加入文件名后加入当前时间
                os.rename(src_file, new_src_file)
        # 移动文件
        shutil.move(src_file, tar_dir)


#获取下载文件地址 timeout延时单位秒
def get_download_file(timeout):  
    download_success = 0
    start_time = time.perf_counter()
    while (True):
        time.sleep(1)
        try:
            lists = os.listdir(download_dir)
            lists.sort(key=lambda fn: os.path.getmtime(download_dir + '\\' + fn))
            download_file = download_dir + '\\' + lists[-1]
            if download_file.endswith('.xlsx') or download_file.endswith('xls') or download_file.endswith('.docx'):
                modify_time = os.path.getmtime(download_file)  # 文件的修改时间
                nowtime = time.time()  # 当前的时间
                if nowtime - modify_time < 120:  # 如果当前时间和文件修改时间间隔小于120秒，判定文件为刚下载的文件
                    download_success = 1
                    break
        except Exception as e:
            showinfo(generate_report.report_status,f'{str(e)}',"black")
            
        end_time = time.perf_counter()
        if end_time - start_time > timeout:
            showinfo(generate_report.report_status,'下载时间超过延时限制',"black")
            break
        else:
            continue
    if download_success == 1:
        return download_file
    else:
        return None

#静默运行
def silent_run():       
    global edge_options
    #global firefox_options
    global driver
    if edge_options is None:
        edge_options = Options()  
    if hide_page.get() == 1:
        edge_options.add_argument("--window-position=0,-2000")
        driver.maximize_window()
    else:
        if driver is not None:
            for arg in edge_options.arguments:
                if arg=="--window-position=0,-2000":
                    arg="--window-position=0,0"
            driver.maximize_window()

#以深度遍历方式寻找文件夹,会迭代寻找path下子文件夹，文件夹数量多时搜索速度很慢    
def search_folder1(path, string):  
    for root, dirs, files in os.walk(path):
        for dir in dirs:
            if string in dir:
                return os.path.join(root, dir)

#以深度遍历方式寻找文件，会迭代寻找path下子文件，文件数量多时搜索速度很慢
def search_file1(path, string):  
    for root, dirs, files in os.walk(path):
        for file in files:
            if string in file:
                return os.path.join(root, file)
            
#搜索path下目录不包括下面的子文件夹，速度快
def search_folder2(path, string): 
    filelist=os.listdir(path)
    for file in filelist:
        if os.path.isdir(path+'\\'+file):
            if string in file:
                return path+'\\'+file
            
#搜索path下目录不包括下面的子文件，速度快            
def search_file2(path,string):
    filelist=os.listdir(path)
    for file in filelist:
        print(file)
        if os.path.isfile(path+'\\'+file):
            if string in file:
                return path+'\\'+file

#如果在path目录下搜不到则搜索次一级目录
def search_folder3(path,string):
    try:
        string1=string.split('.')[0]   #获取单号.前部分  
        string2=string.split('.')[1]   #获取单号.后部分
        len1=len(string1)              #获取序列号长度
        len2=len(string)               #获取测试单号长度
        filelist=os.listdir(path)
        for file in filelist:
            string3=file[0:len1]      
            if string1==string3:
                if os.path.isdir(path+'\\'+file):
                    if string==file[0:len2]:
                        return path+'\\'+file
                    else:                           #在下级目录中寻找
                        filelist1=os.listdir(path+'\\'+file)
                        for file1 in filelist1:
                            if string==file1[0:len2]:                     #os.path.isdir(path+'\\'+file+'\\'+file1) and 
                                return path+'\\'+file+'\\'+file1
    except Exception as e:
        print(str(e))
        
#删除含有特定内容的某个段落
def delete_content(text):
    for paragraph in doc.paragraphs:
        if text in paragraph.text:
            p=paragraph._element
            p.getparent().remove(p)
            paragraph._p=paragraph._element=None
            #doc._body.remove(paragraph._element)  第二种方式删除段落
            break  #仅删除第一个匹配到的段落
          
#使用主页搜索方式下载下载测试规范        
def download_test_spec(testid):
    global driver
    global edge_options
    global max_progress_value
    if isElementExist(alert_button):  #如果存在页面提示框则点击提示框ok
        driver.find_element(By.XPATH,alert_button).click()
        time.sleep(2)
    if driver==None or driver.current_url.split(';')[0]!=laboratories.split(';')[0]:
        initial_browser(laboratories)
        wait_element_clickable((By.XPATH, laboratories_open),'open',30)
        time.sleep(7) #必须
    try:
        showinfo(generate_report.report_status,'定位测试单号输入框...',"black")
        el_input = driver.find_element(By.XPATH,laboratories_input)
        # 获取输入框标签
        showinfo(generate_report.report_status,'清空输入框历史输入记录...',"black")
        el_input.clear()
        showinfo(generate_report.report_status,'输入测试单号...',"black")
        el_input.send_keys(testid)
        showinfo(generate_report.report_status,'提交测试单号...',"black")
        el_input.send_keys('\n')
        time.sleep(2)
        testid1=testid.replace('.','_') 
        testid_item = (By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]')
        wait_element_text(testid_item,'testid_item',testid1,15)
        process=(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div') #进程显示图标
        #等待进程显示图标消失
        wait_element_invisibility(process,'testid process',30)
        time.sleep(2)
        showinfo(generate_report.report_status,'点选查找出来的测试项...',"black")
        el_testid = driver.find_element(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]').click()  #模糊查找
        time.sleep(1)                                     
        showinfo(generate_report.report_status,'等待Quick Infor页面元素刷新显示完毕...',"black")
        quickinfor_testid=(By.XPATH,laboratories_quickinfo)
        wait_element_text((By.XPATH,laboratories_quickinfo),'quickinfo',testid,15)                           
        time.sleep(1)  
        #此时浏览器窗口必须是处于完全平铺,收缩状态则浏览器顶部工具栏里的report按钮不可点击
        showinfo(generate_report.report_status,'点击navigation一栏中的open按钮...',"black")
        driver.find_element(By.XPATH,laboratories_navigation_open).click()
        time.sleep(2)
        showinfo(generate_report.report_status,'点击弹出窗口download按钮...',"black")
        driver.find_element(By.XPATH,laboratories_navigation_open_download).click()  # 点击download
        showinfo(generate_report.report_status,'规范正在下载中...',"black")
        time.sleep(2)
        showinfo(generate_report.report_status,'点击关闭窗口...',"black")
        driver.find_element(By.XPATH,laboratories_navigation_open_download_close).click() 
        time.sleep(2)
    except Exception as e:
        showinfo(generate_report.report_status,'下载测试规范过程中出现异常',"red")

#使用主页搜索方式下载test report至报告文件夹并返回报告地址    
def download_test_report(testid, test_report_path):  
    global driver
    global edge_options
    global max_progress_value
    if isElementExist(alert_button):  #如果存在页面提示框则点击提示框ok
        driver.find_element(By.XPATH,alert_button).click()
        time.sleep(2)
    if driver==None or driver.current_url.split(';')[0]!=laboratories.split(';')[0]:
        initial_browser(laboratories)
        wait_element_clickable((By.XPATH, laboratories_open),'open',30)
        time.sleep(7) #必须
    try:
        showinfo(generate_report.report_status,'定位测试单号输入框...',"black")
        el_input = driver.find_element(By.XPATH,laboratories_input)
        # 获取输入框标签
        showinfo(generate_report.report_status,'清空输入框历史输入记录...',"black")
        el_input.clear()
        showinfo(generate_report.report_status,'输入测试单号...',"black")
        el_input.send_keys(testid)
        showinfo(generate_report.report_status,'提交测试单号...',"black")
        el_input.send_keys('\n')
        time.sleep(2)
        testid1=testid.replace('.','_')
        testid_item = (By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]')
        wait_element_text(testid_item,'testid_item',testid1,15)
        process=(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div') #进程显示图标
        #等待进程显示图标消失
        wait_element_invisibility(process,'testid process',30)
        time.sleep(2)
        showinfo(generate_report.report_status,'点选查找出来的测试项...',"black")
        el_testid = driver.find_element(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]').click()  #模糊查找
        time.sleep(1)                                     
        showinfo(generate_report.report_status,'等待Quick Infor页面元素刷新显示完毕...',"black")
        quickinfor_testid=(By.XPATH,laboratories_quickinfo)
        wait_element_text((By.XPATH,laboratories_quickinfo),'quickinfo',testid,15)                           
        time.sleep(1)  
        #此时浏览器窗口必须是处于完全平铺,收缩状态则浏览器顶部工具栏里的report按钮不可点击
        showinfo(generate_report.report_status,'点击顶部工具栏report按钮...',"black")
        driver.find_element(By.XPATH,laboratories_report).click() #点击顶部工具栏report按钮
        time.sleep(2)
        showinfo(generate_report.report_status,'点击弹出窗口底部Generate按钮...',"black")
        driver.find_element(By.XPATH,laboratories_generate).click()  # 点击Generate
        showinfo(generate_report.report_status,'等待下载报告...',"black")
        time.sleep(5)
    except Exception as e:
        showinfo(generate_report.report_status,'下载测试报告模板中出现异常',"red")
        tkinter.messagebox.showinfo('提示', f'下载测试报告模板中出现异常,请重新点击生成')
        max_progress_value=0
        sys.exit(1)

#执行一个线程任务
def run_thread(target_function,thread_name,*args,**kwargs):
    is_thread_active=0 
    active_threadname=''
    for thread in threading.enumerate():
        if thread.name in active_threadlist and thread.is_alive():
            is_thread_active=1
            active_threadname=thread.name
            break
    if is_thread_active==0:      
        T = threading.Thread(target=target_function,name=thread_name,args=args,kwargs=kwargs)
        T.daemon = True
        T.start()  # 启动线程 
    else:    
        tkinter.messagebox.showinfo('提示', f'{active_threadname}进程正在运行请勿反复点击')



class CustomProgressBar(tkinter.Canvas):
    """自定义进度条控件"""
    def __init__(self, parent, width=400, height=20, **kwargs):
        super().__init__(parent, width=width, height=height, **kwargs)
        self.width = width
        self.height = height
        self.value = 0
        self.max_value = 100
        
        # 创建进度条背景
        self.bg_rect = self.create_rectangle(0, 0, width, height, 
                                            fill="#e0e0e0", outline="#cccccc")
        
        # 创建进度条前景
        self.fg_rect = self.create_rectangle(0, 0, 0, height, 
                                            fill="#4CAF50", outline="#4CAF50")
        
        # 创建文本
        self.text = self.create_text(width/2, height/2, text="0%", 
                                    font=("Arial", 10, "bold"), fill="#333333")
    
    def set_value(self, value):
        """设置进度值 (0-100)"""
        self.value = max(0, min(100, value))
        progress_width = (self.value / 100) * self.width
        
        # 更新进度条宽度
        self.coords(self.fg_rect, 0, 0, progress_width, self.height)
        
        # 更新文本
        self.itemconfig(self.text, text=f"{int(self.value)}%")
        
        # 根据进度改变颜色
        if self.value < 30:
            self.itemconfig(self.fg_rect, fill="#FF5722") # 红色
        elif self.value < 70:
            self.itemconfig(self.fg_rect, fill="#FFC107") # 黄色
        else:
            self.itemconfig(self.fg_rect, fill="#4CAF50") # 绿色

class AdvancedCopyApp:
    def __init__(self, root):
        self.root = root
        # 变量
        self.src_folder = tkinter.StringVar()
        self.dst_folder = tkinter.StringVar()
        self.src_folder.set(raw_data_dir)
        self.dst_folder.set(test_dir)
        self.test_id = tkinter.StringVar()
        self.overwrite = tkinter.IntVar()
        self.overwrite.set(overwrite_current_value)
        self.cobo_team_name=None
        self.is_running = False
        self.total_files = 0
        self.copied_files = 0
        self.total_size = 0
        self.copied_size = 0
        self.setup_ui()
        
    
    def setup_ui(self):
    # 创建主容器框架
        container = ttk.Frame(self.root)
        container.pack(anchor="nw")
        
        # 主框架放在顶部
        main_frame = ttk.Frame(container)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        tool_frame = tkinter.Frame(main_frame)
        tool_frame.pack(fill="x", pady=5)
        checkbutton61 = tkinter.Checkbutton(tool_frame, text='是否覆盖原数据', variable=self.overwrite, onvalue=1, offvalue=0)
        checkbutton61.pack(side="left", fill="x", expand=True, padx=5)
        
        #小组
        #tkinter.Label(tool_frame, text='小组名 ').pack(side="right", fill="x", expand=True, padx=5)
        self.cobo_team_name = ttk.Combobox(tool_frame, width=8, values=teamlist,state="readonly")
        self.cobo_team_name.pack(side="right", expand=True)
        self.cobo_team_name.current(cobo_team_name_current_value)


        # 文件夹选择框架
        folder_frame = ttk.LabelFrame(main_frame, text="文件夹设置")
        folder_frame.pack(fill="x", pady=(0, 20))
        
        # 源文件夹
        src_frame = ttk.Frame(folder_frame)
        src_frame.pack(fill="x")
        # new_button = ttk.Button(folder_frame, text="↺\nClear", width=8)
        # new_button.pack(side="right", expand=True, padx=(10, 0)) 
        ttk.Label(src_frame, text="源文件夹", width=12).pack(side="left")
        src_entry = ttk.Entry(src_frame, textvariable=self.src_folder)
        src_entry.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(src_frame, text="浏览...", command=self.select_source, 
                width=8).pack(side="right")

            
        swap_frame=ttk.Frame(folder_frame)
        swap_frame.pack(fill='x')
        # swap_photo=tk.PhotoImage(file="C:/Users/zhs6cng/OneDrive - Bosch Group/PersonalDrive/python/edataoper/swap.ico")
        ttk.Button(swap_frame, text="swap", command=self.swap_address, 
                width=6).pack(side="top")

    
        # 目标文件夹
        dst_frame = ttk.Frame(folder_frame)
        dst_frame.pack(fill="x")
        
        ttk.Label(dst_frame, text="目标文件夹", width=12).pack(side="left")
        dst_entry = ttk.Entry(dst_frame, textvariable=self.dst_folder)
        dst_entry.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(dst_frame, text="浏览...", command=self.select_destination, 
                width=8).pack(side="right")

             

        # testid框架
        testid_frame = ttk.LabelFrame(main_frame, text="复制单个文件夹", padding=10)
        testid_frame.pack(fill="both", expand=True)
        ttk.Label(testid_frame, text="文件夹名称", width=10).pack(side="left")
        testid_entry = ttk.Entry(testid_frame, textvariable=self.test_id)
        testid_entry.pack(side="left", expand=True)          
        
        # 进度显示框架
        progress_frame = ttk.LabelFrame(main_frame, text="进度", padding=10)
        progress_frame.pack(fill="both", expand=True, pady=(0, 20))
        
        # 自定义进度条
        self.progress_bar = CustomProgressBar(progress_frame, width=500, height=25)
        self.progress_bar.pack(pady=1)
        
        # 进度信息框架
        info_frame = ttk.Frame(progress_frame)
        info_frame.pack(fill="x", pady=1)
        
        # 文件信息
        self.files_label = ttk.Label(info_frame, text="文件: 0/0")
        self.files_label.pack(side="left", padx=20)
        
        # 大小信息
        self.size_label = ttk.Label(info_frame, text="大小: 0 B/0 B")
        self.size_label.pack(side="left", padx=20)
        
        # 速度信息
        self.speed_label = ttk.Label(info_frame, text="速度: 0 B/s")
        self.speed_label.pack(side="right", padx=20)
        
        # 当前文件
        self.current_file_frame = ttk.LabelFrame(progress_frame, text="当前文件", padding=5)
        self.current_file_frame.pack(fill="x", pady=10)
        
        self.current_file_label = ttk.Label(self.current_file_frame, text="等待开始...", 
                                        wraplength=500, anchor="center")
        self.current_file_label.pack()
        
        # 控制按钮框架
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill="x", pady=(0, 10))
        
        # 按钮
        btn_style = {"width": 12, "padding": 5}
        
        self.copy_btn = ttk.Button(control_frame, text="复制", 
                                command=self.start_copy, **btn_style)
        self.copy_btn.pack(side="left", padx=5)
        
        self.move_btn = ttk.Button(control_frame, text="剪切", 
                                command=self.start_move, **btn_style)
        self.move_btn.pack(side="left", padx=5)
        
        self.cancel_btn = ttk.Button(control_frame, text="取消", 
                                    command=self.cancel_operation, 
                                    state="disabled", **btn_style)
        self.cancel_btn.pack(side="left", padx=5)
        
        # 状态栏 - 放在容器底部，不在main_frame内
        status_frame = ttk.Frame(container, relief="sunken", height=25)
        status_frame.pack(side="bottom", fill="x")
        status_frame.pack_propagate(False)
        
        self.status_label = ttk.Label(status_frame, text="就绪", anchor="w")
        self.status_label.pack(side="left", padx=10)
        
        # 时间标签
        self.time_label = ttk.Label(status_frame, text="", anchor="e")
        self.time_label.pack(side="right", padx=10)

        # 启动时间更新
        self.update_time()
    
    def swap_address(self):
        a=self.src_folder.get()
        b=self.dst_folder.get()
        self.src_folder.set(b)
        self.dst_folder.set(a)



    def select_source(self):
        folder = filedialog.askdirectory(title="选择源文件夹")
        if folder:
            self.src_folder.set(folder)
            self.update_status(f"已选择源文件夹: {folder}")
    
    def select_destination(self):
        folder = filedialog.askdirectory(title="选择目标文件夹")
        if folder:
            self.dst_folder.set(folder)
            self.update_status(f"已选择目标文件夹: {folder}")
    
    def update_time(self):
        """更新时间显示"""
        from datetime import datetime
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_label.config(text=current_time)
        self.root.after(1000, self.update_time)
    
    def update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)
    
    def start_copy(self):
        """开始复制"""
        is_thread_active=0 
        active_threadname=''
        for thread in threading.enumerate():
            if thread.name in active_threadlist and thread.is_alive():
                is_thread_active=1
                active_threadname=thread.name
                break

        if not self.validate_inputs():
            return
        
        self.is_running = True
        self.copy_btn.config(state="disabled")
        self.move_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        
        # 在新线程中执行
        if is_thread_active==0: 
            thread = threading.Thread(target=self.copy_folder_with_progress,name="copy")
            thread.daemon = True
            thread.start()
        else:    
            tkinter.messagebox.showinfo('提示', f'{active_threadname}进程正在运行请勿反复点击')    
    
    def start_move(self):
        """开始剪切"""
        is_thread_active=0 
        active_threadname=''
        for thread in threading.enumerate():
            if thread.name in active_threadlist and thread.is_alive():
                is_thread_active=1
                active_threadname=thread.name
                break

        if not self.validate_inputs():
            return
        
        self.is_running = True
        self.copy_btn.config(state="disabled")
        self.move_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        
        # 在新线程中执行
        if is_thread_active==0: 
            thread = threading.Thread(target=self.move_folder_with_progress,name="cut")
            thread.daemon = True
            thread.start()
        else:    
            tkinter.messagebox.showinfo('提示', f'{active_threadname}进程正在运行请勿反复点击')        
    
    def validate_inputs(self):
        """验证输入"""
        src = self.src_folder.get()
        dst = self.dst_folder.get()
        
        if not src or not dst:
            tkinter.messagebox.showerror("错误", "请选择源文件夹和目标文件夹")
            return False
        
        if not os.path.exists(src):
            tkinter.messagebox.showerror("错误", "源文件夹不存在")
            return False
        
        if src == dst:
            tkinter.messagebox.showerror("错误", "源文件夹和目标文件夹不能相同")
            return False
        
        return True
    
    def analyze_folder(self, folder_path):
        """分析文件夹，获取文件总数和总大小"""
        total_files = 0
        total_size = 0
        
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_full_path = os.path.join(root, file)
                if os.path.exists(file_full_path):
                    try:
                        total_size += os.path.getsize(file_full_path)
                        total_files += 1
                    except:
                        pass
        
        return total_files, total_size
    
    def format_size(self, size):
        """格式化文件大小"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.2f} {unit}"
            size /= 1024.0
        return f"{size:.2f} TB"
    
    def copy_folder_with_progress(self):
        """带进度显示的文件夹复制"""
        index=self.overwrite.get()           #是否覆盖原文件
        team_name=self.cobo_team_name.get()  #小组名
        try:
            if self.test_id.get()=='':
                src = Path(self.src_folder.get())
                dst = Path(self.dst_folder.get())    

            elif self.test_id.get()!='' and team_name=='ETC33':
                src= Path(self.src_folder.get()+'/'+self.test_id.get())   
                folder_name = Path(src).name
                dst = Path(self.dst_folder.get()+'/'+folder_name+'/'+folder_name)  #需要在目标文件夹中再生成一个同名文件夹

            else:
                src= Path(self.src_folder.get()+'/'+self.test_id.get())   
                folder_name = Path(src).name
                dst = Path(self.dst_folder.get()+'/'+folder_name)
            
            
            # 分析文件夹
            self.root.after(0, lambda: self.update_status("正在分析文件夹..."))
            self.root.after(0, lambda: self.current_file_label.config(text="正在扫描文件夹..."))
            
            self.total_files, self.total_size = self.analyze_folder(src)
            self.copied_files = 0
            self.copied_size = 0
            
            if self.total_files == 0:
                self.root.after(0, lambda: tkinter.messagebox.showinfo("提示", "源文件夹中没有文件"))
                self.root.after(0, self.reset_buttons)
                return
            
            # 更新UI
            self.root.after(0, self.update_progress_info)
            if index == 1:
                self.root.after(0, lambda: self.update_status(f"开始复制 {self.total_files} 个文件（覆盖模式）"))
            else:
                self.root.after(0, lambda: self.update_status(f"开始复制 {self.total_files} 个文件（不覆盖模式）"))
            
            # 创建目标文件夹
            os.makedirs(dst, exist_ok=True)
            
            import time
            start_time = time.time()
            
            # 复制文件
            for root, dirs, files in os.walk(src):
                if not self.is_running:
                    break
                    
                # 创建子目录
                rel_path = os.path.relpath(root, src)
                target_dir = dst/rel_path
                os.makedirs(target_dir, exist_ok=True)
                
                for file in files:
                    if not self.is_running:
                        break
                    
                    src_file = Path(root) / file
                    
                    if index == 1:
                        # 模式1: 直接覆盖
                        dst_file = target_dir / file
                    else:
                        # 模式0: 检查同名文件，如果存在则修改文件名
                        original_dst_file = target_dir / file
                    
                        if original_dst_file.exists():
                        # 获取文件名和扩展名
                            file_name, file_ext = os.path.splitext(file)
                            
                            # 生成时间戳（格式：YYYYMMDD_HHMMSS）
                            timestamp = time.strftime("%Y%m%d_%H%M%S", time.localtime())      
                            # 新文件名：原文件名_时间戳.扩展名
                            new_file_name = f"{file_name}_{timestamp}{file_ext}"
                            dst_file = target_dir / new_file_name
                            
                            # 显示重命名信息
                            self.root.after(0, lambda f=file, nf=new_file_name: 
                                        self.update_status(f"文件 {f} 重命名为 {nf}"))
                        else:
                            # 没有同名文件，使用原文件名
                            dst_file = original_dst_file
                    try:
                        # 获取文件大小
                        file_size = os.path.getsize(src_file)
                        
                        # 更新UI
                        self.root.after(0, lambda f=file, fs=file_size: 
                                       self.update_current_file(f, fs))
                        
                        # 复制文件
                        shutil.copy2(src_file, dst_file)
                        
                        # 更新进度
                        self.copied_files += 1
                        self.copied_size += file_size
                        
                        # 计算进度
                        if self.total_files > 0:
                            file_progress = (self.copied_files / self.total_files) * 100
                        else:
                            file_progress = 0
                        
                        if self.total_size > 0:
                            size_progress = (self.copied_size / self.total_size) * 100
                        else:
                            size_progress = 0
                        
                        # 使用文件数和大小进度的平均值
                        progress = (file_progress + size_progress) / 2
                        
                        # 计算速度
                        elapsed_time = time.time() - start_time
                        if elapsed_time > 0:
                            speed = self.copied_size / elapsed_time
                        else:
                            speed = 0
                        
                        # 更新UI
                        self.root.after(0, lambda p=progress, s=speed: 
                                       self.update_progress(p, s))
                        
                    except Exception as e:
                        print(f"复制文件 {file} 时出错: {e}")
            
            if self.is_running:
                if index == 1:
                    self.root.after(0, lambda: self.operation_complete("复制（覆盖模式）"))
                else:
                    self.root.after(0, lambda: self.operation_complete("复制（不覆盖模式）"))
            else:
                self.root.after(0, lambda: self.update_status("复制已取消"))
                
        except Exception as e:
            self.root.after(0, lambda: tkinter.messagebox.showerror("错误", f"复制失败: {str(e)}"))
        finally:
            self.root.after(0, self.reset_buttons)
    
    def move_folder_with_progress(self):
        """带进度显示的文件夹移动"""
        # 这里简化为先复制再删除，实际可以使用shutil.move
        # 但为了显示进度，我们还是用复制+删除的方式
        self.copy_folder_with_progress()
        
        # 如果复制成功，删除源文件夹
        if self.is_running and self.copied_files == self.total_files:
            try:
                if self.test_id.get()=='':
                     # 保存文件夹的权限信息
                    stat_info = os.stat(self.src_folder.get())
                    shutil.rmtree(self.src_folder.get())    
                     # 等待一下确保删除完成
                    time.sleep(0.1)
                    # 重新创建文件夹
                    os.makedirs(self.src_folder.get(), exist_ok=True)
                    # 恢复原始权限
                    os.chmod(self.src_folder.get(), stat_info.st_mode)
                else:
                    shutil.rmtree(self.src_folder.get()+'/'+self.test_id.get())    
                self.root.after(0, lambda: self.update_status("移动完成"))
            except Exception as e:
                self.root.after(0, lambda: tkinter.messagebox.showwarning("警告", 
                    f"复制完成，但删除源文件夹失败: {str(e)}"))
    
    def update_current_file(self, filename, filesize):
        """更新当前文件显示"""
        size_str = self.format_size(filesize)
        self.current_file_label.config(text=f"{filename} ({size_str})")
    
    def update_progress(self, value, speed=0):
        """更新进度显示"""
        # 更新进度条
        self.progress_bar.set_value(value)
        
        # 更新文件计数
        self.files_label.config(text=f"文件: {self.copied_files}/{self.total_files}")
        
        # 更新大小信息
        copied_size_str = self.format_size(self.copied_size)
        total_size_str = self.format_size(self.total_size)
        self.size_label.config(text=f"大小: {copied_size_str}/{total_size_str}")
        
        # 更新速度
        speed_str = self.format_size(speed) + "/s" if speed > 0 else "0 B/s"
        self.speed_label.config(text=f"速度: {speed_str}")
    
    def update_progress_info(self):
        """初始化进度信息"""
        self.files_label.config(text=f"文件: 0/{self.total_files}")
        total_size_str = self.format_size(self.total_size)
        self.size_label.config(text=f"大小: 0 B/{total_size_str}")
        self.speed_label.config(text="速度: 0 B/s")
        self.progress_bar.set_value(0)
    
    def operation_complete(self, operation):
        """操作完成"""
        self.update_status(f"{operation}完成!")
        self.progress_bar.set_value(100)
        tkinter.messagebox.showinfo("完成", f"文件夹{operation}完成!")
        self.current_file_label.config(text=f"{operation}完成!")
    
    def cancel_operation(self):
        """取消操作"""
        self.is_running = False
        self.cancel_btn.config(state="disabled")
        self.update_status("正在取消...")
    
    def reset_buttons(self):
        """重置按钮状态"""
        self.is_running = False
        self.copy_btn.config(state="normal")
        self.move_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        self.update_status("就绪")

class GenerateTestReport:
    def __init__(self, root):
        self.root = root
        # 变量
        self.hide_page=tkinter.IntVar()
        self.hide_page.set(hide_page_current_value)
        self.compress_picture = tkinter.IntVar()                  #压缩图片勾选
        self.compress_picture.set(compress_picture_current_value)
        self.time_sort = tkinter.IntVar()                  #时间排序勾选
        self.time_sort.set(time_sort_current_value)
        self.curve_generate = tkinter.IntVar()                  #生成曲线勾选
        self.curve_generate.set(curve_generate_current_value)    
        self.testid_input = tkinter.StringVar()            #test id
        self.link_title_input = tkinter.StringVar()            #link title
        self.data_title_input = tkinter.StringVar()            #Data titles
        self.need_report=tkinter.IntVar()                      #是否需要报告
        self.need_report.set(need_report_current_value)   
        self.first_approval=None
        self.second_approval=None
        self.equipment_select=None
        self.pre_conclusion_select=None
        self.post_conclusion_select=None
        self.report_status=None                          #生成报告界面进度显示栏
        self.posttest_frame=None

        self.lab_temperature = tkinter.StringVar() 
        self.lab_humidity = tkinter.StringVar() 
        self.test_status=None                              #test status

        self.setup_ui()

    def on_equipment_changed(self,*args):
        self.equipment_select['values'] = equipment_list

    def on_approval_changed(self,*args): 
        self.first_approval['values']=approval_list
        self.first_approval['values']=approval_list

    def on_conclusion_changed(self,*args):  
        self.pre_conclusion_select['values']=conclusion_list
        self.post_conclusion_select['values']=conclusion_list

    def on_combobox35_select(self,event):
        selected=self.test_status.get()
        if selected=='IN_PROGRESS':
            data_range_picker=DateRangePicker(window) 


    def setup_ui(self):
    # 创建主容器框架
        tool_frame = tkinter.Frame(self.root)
        tool_frame.grid(row=1, column=1,stick=W)
        tkinter.Checkbutton(tool_frame, text='隐藏抓取页面',variable=self.hide_page, onvalue=1, offvalue=0).grid(row=1, column=1, sticky=W)
        tkinter.Checkbutton(tool_frame, text='压缩图片', variable=self.compress_picture, onvalue=1, offvalue=0).grid(row=1, column=2, sticky=W)
        tkinter.Checkbutton(tool_frame, text='时间排序', variable=self.time_sort, onvalue=1, offvalue=0).grid(row=1, column=3, sticky=W)
        tkinter.Checkbutton(tool_frame, text='生成曲线', variable=self.curve_generate, onvalue=1, offvalue=0).grid(row=1, column=4, sticky=W)
        #report_frame
        report_frame = tkinter.LabelFrame(self.root, text='报告基础信息', width='2600', height='800',font='Arial 12 bold italic')
        report_frame.grid(row=2, column=1,columnspan=2,sticky=W)
        tkinter.Label(report_frame, text='输入test ID ').grid(row=1, column=1, sticky=W)
        tkinter.Entry(report_frame, width=16, textvariable=self.testid_input).grid(row=1, column=2,columnspan=2, sticky=W)  # test id 输入框

        
        tkinter.Label(report_frame, text='更改test status ').grid(row=1, column=5, sticky=W)
        self.test_status = ttk.Combobox(report_frame, width=20, values=test_status_list,state="readonly")
        #self.test_status.bind("<<ComboboxSelected>>",self.on_combobox35_select)
        self.test_status.grid(row=1, column=6,sticky=W)
        ttk.Button(report_frame, text='确认', width=8, command=lambda:run_thread(generate_report.update_test_status,'update_test_status')).grid(row=1, column=7, padx=5, sticky=W)


        ttk.Button(report_frame, text='生成报告', width=8, command=lambda:run_thread(generate_report.generate_test_report,'generate_test_report')).grid(row=1, column=3, padx=5, sticky=W)
        ttk.Button(report_frame, text='上传报告', width=8, command=lambda:run_thread(generate_report.upload_test_report,'upload_test_report')).grid(row=1, column=4, padx=5, sticky=W)
        tkinter.Label(report_frame, text='输入title ').grid(row=2, column=1, sticky=W)
        tkinter.Entry(report_frame, width=16, textvariable=self.link_title_input).grid(row=2, column=2, sticky=W)  # title 输入框
        ttk.Button(report_frame, text='上传link', width=8, command=lambda:run_thread(generate_report.upload_link,'upload_link')).grid(row=2, column=3, padx=5, sticky=W)

        ttk.Button(report_frame, text='更新规范', width=8, command=lambda:run_thread(generate_report.update_test_spec,'update_test_spec')).grid(row=2, column=4, padx=5, sticky=W)

        tkinter.Label(report_frame, text='输入title ').grid(row=3, column=1, sticky=W)
        
        tkinter.Entry(report_frame, width=16, textvariable=self.data_title_input).grid(row=3, column=2, sticky=W)  # title 输入框
        ttk.Button(report_frame, text='上传data', width=8, command=lambda:run_thread(generate_report.upload_data,'upload_data')).grid(row=3, column=3, padx=5, sticky=W)
        tkinter.Checkbutton(report_frame, text='是否需要报告', variable=self.need_report, onvalue=1, offvalue=0).grid(row=3, column=4, padx=5,sticky=W)


        tkinter.Label(report_frame, text='添加设备').grid(row=4, column=1, sticky=W) 
        self.equipment_select = Combopicker(report_frame, entrywidth='60', values=equipment_list)
        self.equipment_select.grid(row=4, column=2,columnspan=4)
        #ttk.Button(report_frame, text='清空', width=4, command=lambda:clear_input(self.equipment_select)).grid(row=4, column=6, padx=5, sticky=W)
        tkinter.Label(report_frame, text='结论描述(前)').grid(row=5, column=1,sticky=W)
        self.pre_conclusion_select = Combopicker(report_frame, entrywidth='60', values=conclusion_list)
        self.pre_conclusion_select.grid(row=5, column=2,columnspan=4)
        #ttk.Button(report_frame, text='清空', width=4, command=lambda:clear_input(self.pre_conclusion_select)).grid(row=5, column=6, padx=5, sticky=W)
        tkinter.Label(report_frame, text='结论描述(后)').grid(row=6, column=1,sticky=W)
        self.post_conclusion_select = Combopicker(report_frame, entrywidth='60', values=conclusion_list)
        self.post_conclusion_select.grid(row=6, column=2,columnspan=4)
        #ttk.Button(report_frame, text='清空', width=4, command=lambda:clear_input(self.post_conclusion_select)).grid(row=6, column=6, padx=5, sticky=W)
        tkinter.Label(report_frame, text='审批人1', width=12).grid(row=7, column=1, stick=NW)
        self.first_approval = ttk.Combobox(report_frame, width=27, values=approval_list)
        if first_approval_current_value != -1:
            self.first_approval.current(first_approval_current_value)
        self.first_approval.grid(row=7, column=2, columnspan=2, sticky=NW)
        tkinter.Label(report_frame, text='审批人2', width=12).grid(row=8, column=1, stick=NW)
        self.second_approval = ttk.Combobox(report_frame, width=27, values=approval_list)
        if second_approval_current_value != -1:
            self.second_approval.current(second_approval_current_value)
        self.second_approval.grid(row=8, column=2, columnspan=2, sticky=NW)


        tkinter.Label(report_frame, text='实验室温度', width=8).grid(row=7, column=4, stick=W)
        tkinter.Entry(report_frame, width=8, textvariable=self.lab_temperature).grid(row=7, column=5, sticky=W)
        tkinter.Label(report_frame, text='实验室湿度', width=8).grid(row=8, column=4, stick=W)
        tkinter.Entry(report_frame, width=8, textvariable=self.lab_humidity).grid(row=8, column=5, sticky=W)


        self.posttest_frame = tkinter.LabelFrame(tab1, width='200', height='200', text='后测试',font='Arial 12 bold italic')
        self.posttest_frame.grid(row=3, column=2,columnspan=2, sticky=NW)
        ttk.Button(self.posttest_frame, text='添加', width='6', command=self.add_post_test).grid(row=1, column=2)
        ttk.Button(self.posttest_frame, text='删除', width='6', command=self.delete_post_test).grid(row=1, column=3,columnspan=2,sticky=W)
        tkinter.Label(self.posttest_frame, text='后测试ID', width=12).grid(row=2, column=1)
        tkinter.Label(self.posttest_frame, text='后测试名', width=12).grid(row=2, column=2)
        tkinter.Label(self.posttest_frame, text='后测试结果', width=13).grid(row=2, column=3, columnspan=2)
    
        # status_frame
        status_frame = tkinter.LabelFrame(self.root, text='进度显示',font='Arial 12 bold italic')
        status_frame.grid(row=3, column=1, sticky=NW)
        self.report_status = scrolledtext.ScrolledText(status_frame, width=50, height=9)  # 进 度显示栏
        self.report_status.tag_config("green",foreground="green")
        self.report_status.tag_config("red",foreground="red")
        self.report_status.tag_config("black",foreground="black")
        self.report_status.grid(row=1, column=1, columnspan=3,sticky=W)
        ttk.Button(status_frame, text='清空状态信息', command=lambda:clear_status(self.report_status)).grid(row=2, column=1, sticky=W)
        progressbar=ttk.Progressbar(status_frame,orient='horizontal',length=405,mode='determinate')
        progressbar["maximum"]=max_progress_value
        progressbar["value"]=current_progress
        progressbar.grid(row=3,column=1,sticky=W)
        button_function = Button(status_frame, text='测试功能', command=show_current)
        button_function.grid(row=4,column=1,sticky=W)
    
        #生成报告
    def generate_test_report(self):
        global test_id
        global doc
        global report_dir
        global excel_file_list
        global current_progress
        global max_progress_value

        #参数初始化,必须
        report_dir=''
        win_word = None
        word_doc = None
        win_excel= None
        excel_file_list.clear()
        max_progress_value=100
        current_progress=0
        #update_progress()

        # 获取测试报告地址
        showinfo(self.report_status,'获取测试报告地址...',"black")
        report_dir,test_data_path=self.get_report_dir()
        try:
            doc = Document(report_dir)
        except Exception as e:
            showinfo(self.report_status,'报告模板打开失败',"red")
            showinfo(self.report_status,str(e),"red")
            tkinter.messagebox.showinfo('提示', '报告模板打开失败,请手动sava as报告模板再重试')
            sys.exit(1)

        current_progress=10
        if(self.curve_generate.get()==1):
            #搜寻数据文件夹下所有的测试原始数据并生成曲线
            showinfo(self.report_status,'原始数据生成曲线中...',"black")
            test_curve.generate_curve(test_data_path,1)
            
            showinfo(self.report_status,'曲线生成完毕...',"black")    

        #插入附件和excel数据
        showinfo(self.report_status,'插入附件或excel数据...',"black")
        self.insert_document(test_data_path)
        current_progress=30

        # #先删除报告模板规范中自带的图片,如果不提前删除则后续插入测试数据图片会失败
        showinfo(self.report_status,'删除模板里规范部分自带图片中...',"black")
        doc = Document(report_dir) #此处必须再定义一次，否则附件和excel数据都会插入失败,因为使用win32com也打开了doc文档并且win32com打开的word文档进行了保存，如果这里不定义则最后的save会覆盖win32com对doc的修改
        image_info = self.save_image_info(doc)
        self.delete_inline_images(doc)
        current_progress = 40    

        # 插入设备信息
        showinfo(self.report_status,'插入设备信息中...',"black")
        self.insert_equipment()
        showinfo(self.report_status,'设备信息插入完毕',"green")
        current_progress = 50   

        #插入结论和审批人员
        showinfo(self.report_status,'插入结论审批人信息中...',"black")
        self.add_conclusion_approval()
        showinfo(self.report_status,'结论审批人信息插入完毕',"green")
        current_progress = 60

        #插入图片
        showinfo(self.report_status,'插入图片中...',"black")
        self.insert_data(test_data_path)
        showinfo(self.report_status,'图片数据插入完毕',"green")
        current_progress = 90 

        #所有数据插入完毕以后恢复测试规范中的图片
        showinfo(self.report_status,'恢复模板里规范部分自带图片中...',"black")
        self.restore_images(doc, image_info)
        current_progress = 95

        #将末尾的excel表格数据移动到指定段落
        showinfo(self.report_status,'移动excel表格数据到指定段落中...',"black")
        tablelist = doc.tables
        paragraphs = doc.paragraphs  #这里需要重新生成一下
        for i in range(len(excel_file_list)):
            excel_file_name=os.path.basename(excel_file_list[i][2]).split('.')[0]
            for paragraph in paragraphs:
                if paragraph.text==excel_file_name:
                    self.move_table(tablelist[len(tablelist)-i-1],paragraph)
        current_progress = 97 

        #在Test results and evaluation处增加结论描述
        conclusions = self.post_conclusion_select.current_value
        if conclusions!='':
            conclusions_list = conclusions.split('&')
            paragraphs = doc.paragraphs  #这里需要重新生成一下
            for i in range(len(paragraphs)):
                if paragraphs[i].text=="Test results and evaluation":
                    p=paragraphs[i+1].insert_paragraph_before()
                    for conclusion in conclusions_list:
                        run = p.add_run(conclusion + '\n')
                        run.font.size = Pt(12) 

        #在报告最末尾加上结束语           
        text='................................................................................End of Test Report......................................................................................'  
        doc.add_paragraph(text)


        #保存报告
        showinfo(self.report_status,'保存报告中...',"black")
        try:
            doc.save(report_dir)
        except PermissionError as e:
            showinfo(self.report_status,'测试报告被打开无法被保存,请手动关闭文件再重试一次',"red")
            tkinter.messagebox.showinfo('提示', '测试报告被打开无法被保存,请手动关闭文件再重试一次')
        else:
            showinfo(self.report_status,f'{test_id}报告生成完毕',"green")
            current_progress = 100
            result=tkinter.messagebox.askyesno("hello","报告生成完毕,是否立即打开报告?")
            report_dir1='""'+" "+'"'+report_dir+'"'  #此处为了能使得os.system调用dos命令必须对地址进行相应处理 否则dos命令无法识别
            if(result):
                os.system(f"start {report_dir1}")

        #移动表格至指定段落    
    def move_table(self,table,paragraph):  
        table._element.getparent().remove(table._element)
        paragraph._element.addnext(table._tbl)


    #删除表格多余单元格边框
    def delete_cell_border(self,table,column,row,excess_cell_number):
        global doc
        cell1=None
        cell2=None
        #只能按行删除不能按列删除，否则报index out of list错误，可能原因:删除某行某个单元格会影响到同一行别的单元格
        for i in range(excess_cell_number):   #删除倒数第一行各个单元格
            cell1=table.cell(row-1,column-1-i)
            cells1=cell1._tc
            row_cells1=cell1._element.getparent()
            row_cells1.remove(cells1)
            
        for j in range(excess_cell_number):   #删除倒数第二行各个单元格
            cell2=table.cell(row-2,column-1-j)
            cells2=cell2._tc
            row_cells2=cell2._element.getparent()
            row_cells2.remove(cells2)

    #根据测试id获取report地址
    def get_report_dir(self):
        global max_progress_value
        global report_dir
        test_id = self.testid_input.get()
        if test_id=='debug':
            button_function.grid(row=4,column=1,sticky=W)
            max_progress_value=0
            sys.exit(1)
        if test_id=='':
            tkinter.messagebox.showinfo(self.report_status,'提示', '输入空值无效')
            max_progress_value=0
            sys.exit(1)
        showinfo(self.report_status,'正在搜索测试文件夹中....',"black")
        test_folder_path = search_folder3(test_dir, test_id)
        if test_folder_path is None:
            showinfo(self.report_status,'测试文件夹不存在',"red")
            tkinter.messagebox.showinfo('提示', '测试文件夹不存在')
            max_progress_value=0
            sys.exit(1)
        else:
            test_data_path = search_folder1(test_folder_path, "Test Data")
            test_report_path = search_folder1(test_folder_path, "Test Report")
            if test_data_path is None:
                showinfo(self.report_status,'数据文件夹不存在',"red")
                tkinter.messagebox.showinfo('提示', '数据文件夹不存在')
                max_progress_value=0
                sys.exit(1)
            if test_report_path is None:
                showinfo(self.report_status,'存放报告文件夹不存在',"red")
                tkinter.messagebox.showinfo('提示', '存放报告文件夹不存在')
                max_progress_value=0
                sys.exit(1)
            else:
                file_list = os.listdir(test_report_path)
                for file in file_list:  # 在报告文件夹中搜索是否存在报告模板，如果无则去主页下载模板
                    if file.startswith('TS'):
                        report_dir = test_report_path + '\\' + file
                        break
                if report_dir == '':
                    showinfo(self.report_status,'下载测试报告中...',"black")
                    download_test_report(test_id, test_report_path)
                    download_file = get_download_file(15)
                    if download_file is not None:
                        filename = os.path.basename(download_file)
                        showinfo(self.report_status,'移动测试报告模板中...',"black")
                        move_file(download_file, test_report_path,0)    
                        report_dir=test_report_path + '\\' + filename
                        showinfo(self.report_status,f'{test_id}测试报告模板下载成功',"green")
                    else:
                        showinfo(self.report_status,f'超时,{test_id}测试报告模板下载失败,请检查网络或其他原因',"red")
                        tkinter.messagebox.showinfo('提示', '报告测试模板下载失败,请检查网络或其他原因')
                        max_progress_value=0
                        sys.exit(1)     
                return report_dir,test_data_path


    #删除报告模板规范部分自带图片        
    def delete_inline_images(self,doc): 
        try:
            for paragraph in doc.paragraphs:
                for run in paragraph.runs:
                    for inline_shape in run.element.xpath('.//w:drawing'):
                        inline_shape.getparent().remove(inline_shape)
        except Exception as e:
            showinfo(generate_report.report_status,f'删除报告模板规范部分自带图片时报错:{str(e)}','red')                   
        table_paragraphlist=[]
        table_imglist=[]
        table_picture_list=[]
        #删除报告模板中表格单元格内的图片
        try:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        #删除表格单元格内嵌的表格
                        for element in cell._element.xpath('.//w:tbl'):
                            element.getparent().remove(element)
                        #删除表格单元格内嵌的图片
                        for p in cell.paragraphs:
                            if p._element.xpath('.//w:drawing'):
                                table_paragraphlist.append(p)
                                table_picture_list = self.get_all_picture(doc, p)  # 获取段落表格单元格包含的所有图片                        
                                table_imglist.append(table_picture_list)
                                p.clear()
        except Exception as e:
            showinfo(generate_report.report_status,f'删除表格单元格内嵌图片时报错:{str(e)}','red')                        
                
    #保存报告模板中自带图片                              
    def save_image_info(self,doc):
        image_info = []
        try:
            for paragraph in doc.paragraphs:
                for run in paragraph.runs:
                    # 查找内联图片
                    for inline_shape in run.element.xpath('.//w:drawing'):
                        # 查找图片的 Blip 元素
                        blip_elements = inline_shape.xpath('.//a:blip')
                        if blip_elements:
                            blip = blip_elements[0]
                            # 获取图片的嵌入 ID
                            image_embed_id = blip.get(qn('r:embed'))
                            if image_embed_id:
                                # 获取图片的宽度和高度
                                extent = inline_shape.xpath('.//wp:extent')
                                if extent:
                                    width = int(extent[0].get('cx', 0))  # 宽度（EMU 单位）
                                    height = int(extent[0].get('cy', 0))  # 高度（EMU 单位）
                                    # 保存图片信息
                                    image_info.append({
                                        'paragraph': paragraph,
                                        'run': run,
                                        'image_embed_id': image_embed_id,
                                        'width': width,
                                        'height': height
                                    })
        except Exception as e:
            showinfo(generate_report.report_status,f'保存报告模板自带图片时报错:{str(e)}','red')
        return image_info

    #恢复报告模板中自带图片   
    def restore_images(self,doc, image_info):
        for info in image_info:
            try:
                paragraph = info['paragraph']
                run = info['run']
                image_embed_id = info['image_embed_id']
                width = info['width'] / 914400  # 转换为 Inches
                height = info['height'] / 914400  # 转换为 Inches
                # 查找图片的二进制数据
                image_part = doc.part.related_parts[image_embed_id]
                with open("temp_image.png", "wb") as f:
                    f.write(image_part.blob)
                # 重新插入图片
                run.add_picture("temp_image.png", width=Inches(width), height=Inches(height))
            except Exception as e:
                showinfo(generate_report.report_status,f'恢复图片时发现无法识别的图片格式:{str(e)}','red')

    #获取段落中所有图片        
    def get_all_picture(self,document: Document, paragraph: Paragraph):
        try:
            all_img = paragraph._element.xpath('.//pic:pic')
            if not all_img:
                return
            size = len(all_img)
            allimage = []
            for i in range(size):
                img: CT_Picture = all_img[i]
                embed = img.xpath('.//a:blip/@r:embed')[0]
                related_part: ImagePart = document.part.related_parts[embed]
                image: Image = related_part.image
                print("image" + str(i) + ':' + str(image.width) + '-' + str(image.px_width) + '-' + str(image.horz_dpi))
                allimage.append(image)
            return allimage
        except UnrecognizedImageError as e:
            showinfo(generate_report.report_status,f'异常描述:在测试规范条件中发现无法识别的图片格式:{str(e)}','red')
            return allimage


    #压缩图片
    def compress_picture_size(self,picture_path, quality):
        picture_name = os.path.basename(picture_path)
        picture_dir = os.path.dirname(picture_path)
        # if not os.path.exists(picture_dir + '\compressed'):
        #     os.mkdir(picture_dir + '\compressed')
        # compressed_picture_path = picture_dir + '\compressed\\' + picture_name
        compressed_dir = os.path.join(picture_dir, 'compressed')
        os.makedirs(compressed_dir, exist_ok=True)
        compressed_picture_path = os.path.join(compressed_dir, picture_name)


        picture = Image.open(picture_path)
        picture.save(compressed_picture_path, quality=quality)
        return compressed_picture_path


    #上传测试报告        
    def upload_test_report(self):
        global driver
        global edge_options
        global max_progress_value
        global test_dir
        report_dir=''
        testid = self.testid_input.get()
        if testid=='':
            tkinter.messagebox.showinfo('提示', '输入空值无效')
            max_progress_value=0
        else:    
            showinfo(generate_report.report_status,'正在搜索测试文件夹中....',"black") 
            test_folder_path=search_folder3(test_dir,testid)
            if test_folder_path is None:
                showinfo(generate_report.report_status,'测试文件夹不存在',"red")
                tkinter.messagebox.showinfo('提示', '测试文件夹不存在')
                max_progress_value=0
                sys.exit(1)
            else:
                test_report_path = search_folder1(test_folder_path, "Test Report")
                if test_report_path is None:
                    showinfo(generate_report.report_status,'存放报告文件夹不存在',"red")
                    tkinter.messagebox.showinfo('提示', '存放报告文件夹不存在')
                    max_progress_value=0
                    sys.exit(1)
                else:
                    file_list = os.listdir(test_report_path)
                    for file in file_list:  # 在报告文件夹中搜索是否存在报告
                        if file.startswith('TS'):
                            report_dir = test_report_path + '\\' + file
                            break
                    if report_dir == '':
                        showinfo(generate_report.report_status,'报告文件夹内无报告文件,无法上传',"red")
                        tkinter.messagebox.showinfo('提示', '报告文件夹内无报告文件,无法上传')
                        max_progress_value=0
                    else:
                        try:        
                            if driver is None or 'laboratories' not in driver.current_url:
                                initial_browser(laboratories)
                                #driver.maximize_window()     
                                wait_element_clickable((By.XPATH,laboratories_open),'open',30)
                                time.sleep(7) #必须
                                showinfo(generate_report.report_status,'定位测试单号输入框...',"black")
                                el_input = driver.find_element(By.XPATH,laboratories_input)
                                # 获取输入框标签
                                showinfo(generate_report.report_status,'清空输入框历史输入记录...',"black")
                                el_input.clear()
                                showinfo(generate_report.report_status,'输入测试单号...',"black")
                                el_input.send_keys(testid)
                                showinfo(generate_report.report_status,'提交测试单号...',"black")
                                el_input.send_keys('\n')
                                time.sleep(2)                     
                                testid1=testid.replace('.','_')
                                testid_item = (By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]')
                                wait_element_text(testid_item,'testid_item',testid1,15)
                                process=(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div') #进程显示图标
                                #等待进程显示图标消失
                                wait_element_invisibility(process,'testid process',30)
                                showinfo(generate_report.report_status,'点选查找出来的测试项...',"black")
                                el_testid = driver.find_element(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]').click()  #模糊查找            
                                quickinfor_testid=(By.XPATH,laboratories_quickinfo)
                                wait_element_text((By.XPATH,laboratories_quickinfo),'quickinfo',testid,15)  
                                time.sleep(4)  
                            
                            #此处做一下判断:el_quickinfor中的testid是否是与当前程序输入框中(非edata页面中的输入框)testid一致,如果不一致则重新在edata页面中的输入框中输入程序输入框中的testid
                            el_quickinfor_testid=driver.find_element(By.XPATH,laboratories_quickinfo) 
                            if isElementExist(alert_button):  #如果存在页面提示框则点击提示框ok
                                driver.find_element(By.XPATH,alert_button).click()
                                time.sleep(2)          
                            if el_quickinfor_testid.text.strip()!=testid:
                                showinfo(generate_report.report_status,'清空输入框历史输入记录...',"black")
                                el_input = driver.find_element(By.XPATH,laboratories_input)
                                el_input.clear()
                                time.sleep(1)
                                showinfo(generate_report.report_status,'输入测试单号...',"black")
                                el_input.send_keys(self.testid_input.get())
                                showinfo(generate_report.report_status,'提交测试单号...',"black")
                                el_input.send_keys('\n')
                                time.sleep(2)
                                testid1=self.testid_input.get().replace('.','_')
                                testid_item = (By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]')
                                wait_element_text(testid_item,'testid_item',testid1,15)
                                process=(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div') #进程显示图标
                                #等待进程显示图标消失
                                wait_element_invisibility(process,'testid_process',30)
                                showinfo(generate_report.report_status,'点选查找出来的测试项...',"black")
                                el_testid = driver.find_element(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]').click()  #模糊查找  
                                quickinfor_testid=(By.XPATH,laboratories_quickinfo)
                                wait_element_text((By.XPATH,laboratories_quickinfo),'quickinfo',testid,15)                         
                                time.sleep(4)  
                    
                            #点击documents    
                            showinfo(generate_report.report_status,'点击document',"black")
                            if isElementExist(alert_button):  #如果存在页面提示框则点击提示框ok
                                driver.find_element(By.XPATH,alert_button).click()
                            time.sleep(2)
                            driver.find_element(By.XPATH,laboratories_documents).click()
                            #点击new
                            wait_element_clickable((By.XPATH,laboratories_new),'new',15)
                            time.sleep(2)
                            driver.find_element(By.XPATH,laboratories_new).click()
                            #点击Upload Document
                            showinfo(generate_report.report_status,'点击Upload Document',"black")
                            time.sleep(2)
                            driver.find_element(By.XPATH,laboratories_upload_document_img).click()
                            time.sleep(1)
                            #输入报告title
                            showinfo(generate_report.report_status,'输入报告title',"black")
                            title_input=driver.find_element(By.XPATH,laboratories_title)
                            title_input.clear()
                            time.sleep(1)
                            report_name=os.path.basename(report_dir)
                            title=os.path.splitext(report_name)[0]
                            title_input.send_keys(title)
                            time.sleep(1)
                            #提交报告地址
                            showinfo(generate_report.report_status,'提交报告地址',"black")
                            driver.find_element(By.XPATH,laboratories_add_document).send_keys(report_dir)
                            time.sleep(2)
                            #点击save保存
                            showinfo(generate_report.report_status,'点击save',"black")
                            driver.find_element(By.XPATH,laboratories_save).click()
                            time.sleep(2)
                            showinfo(generate_report.report_status,'报告上传完毕',"green")
                            tkinter.messagebox.showinfo('提示', f'{testid}报告上传完毕')
                        except Exception as e:
                            showinfo(generate_report.report_status,'上传测试报告模板中出现异常',"red")
                            tkinter.messagebox.showinfo('提示', f'上传测试报告模板中出现异常,请重新点击上传')
                            max_progress_value=0

    #上传link
    def upload_link(self):
        global test_dir
        title = self.link_title_input.get()
        testid= self.testid_input.get()
        link_dir=''
        if title=='':
            tkinter.messagebox.showinfo('提示', '输入空值无效')
            max_progress_value=0
            sys.exit(1)
        showinfo(generate_report.report_status,'正在搜索测试文件夹中....',"black") 
        test_folder_path=search_folder3(test_dir,testid)
        if test_folder_path is None:
            showinfo(generate_report.report_status,'测试文件夹不存在',"red")
            tkinter.messagebox.showinfo('提示', '测试文件夹不存在')
            max_progress_value=0
            sys.exit(1)
        else:
            file_list = os.listdir(test_folder_path)
            for file in file_list:  # 在测试文件夹中搜索是否存在link 文件
                if file.startswith('Raw data file link'):
                    link_dir = test_folder_path + '\\' + file
                    break
            if link_dir == '':
                showinfo(generate_report.report_status,'测试文件夹内无link文件,正在创建link文件中....',"red")
                #创建link文件
                doc_address=Document()
                paragraph=doc_address.add_paragraph()
                run=paragraph.add_run()
                unc_path=test_dir_driver_letter_unc_path+'\\'+test_folder_path.split(':',1)[1]
                add_hyperlink(paragraph,'ctr+鼠左键点击这里访问测试数据',unc_path)
                doc_address.save(test_folder_path+'\\Raw data file link.docx')
                link_dir=test_folder_path+'\\Raw data file link.docx'
                showinfo(generate_report.report_status,'link文件创建完毕',"green")
            self.upload_file(testid,title,link_dir)   


    #上传数据
    def upload_data(self):
        global test_dir
        global needreport
        needreport=0 
        title = self.data_title_input.get()
        testid= self.testid_input.get()
        data_dir=''
        if title=='':
            tkinter.messagebox.showinfo('提示', '输入空值无效')
            max_progress_value=0
            sys.exit(1)
        showinfo(generate_report.report_status,'正在搜索测试文件夹中....',"black") 
        test_folder_path=search_folder3(test_dir,testid)
        if test_folder_path is None:
            showinfo(generate_report.report_status,'测试文件夹不存在',"red")
            tkinter.messagebox.showinfo('提示', '测试文件夹不存在')
            max_progress_value=0
            sys.exit(1)
        else:
            folder_path=test_folder_path+'\\Original data'
            showinfo(generate_report.report_status,'正在压缩原始数据中...',"black")
            data_dir=zip_folder(folder_path, test_folder_path)
            showinfo(generate_report.report_status,'压缩完毕，正在上传数据中...',"black")
            self.upload_file(testid,title,str(data_dir))             #data_dir是Path类型对象，必须转换为字符串

    #上传数据             
    def upload_file(self,testid,title,file_dir):
        global driver
        global edge_options
        global max_progress_value
        try:        
            if driver is None or 'laboratories' not in driver.current_url:
                initial_browser(laboratories)
                #driver.maximize_window()     
                wait_element_clickable((By.XPATH,laboratories_open),'open',30)
                time.sleep(7) #必须
                showinfo(generate_report.report_status,'定位测试单号输入框...',"black")
                el_input = driver.find_element(By.XPATH,laboratories_input)
                # 获取输入框标签
                showinfo(generate_report.report_status,'清空输入框历史输入记录...',"black")
                el_input.clear()
                showinfo(generate_report.report_status,'输入测试单号...',"black")
                el_input.send_keys(testid)
                showinfo(generate_report.report_status,'提交测试单号...',"black")
                el_input.send_keys('\n')
                time.sleep(2)                     
                testid1=testid.replace('.','_')

                testid_item = (By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]')
                wait_element_text(testid_item,'testid_item',testid1,15)
                process=(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div') #进程显示图标
                #等待进程显示图标消失
                wait_element_invisibility(process,'testid process',30)
                showinfo(generate_report.report_status,'点选查找出来的测试项...',"black")
                el_testid = driver.find_element(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]').click()  #模糊查找   
                quickinfor_testid=(By.XPATH,laboratories_quickinfo)
                wait_element_text((By.XPATH,laboratories_quickinfo),'quickinfo',testid,15)  
                #如果输入框输入的testid后缀是'_',需要将_变成. 因为quickinfor的testid是.
                wait_element_text(quickinfor_testid,'quick infor',testid,15)                 
                time.sleep(4)  
            if isElementExist(alert_button):  #如果存在页面提示框则点击提示框ok
                driver.find_element(By.XPATH,alert_button).click()
                time.sleep(2)
            #此处做一下判断:el_quickinfor中的testid是否是与当前程序输入框中(非edata页面中的输入框)testid一致,如果不一致则重新在edata页面中的输入框中输入程序输入框中的testid
            el_quickinfor_testid=driver.find_element(By.XPATH,laboratories_quickinfo)        
            if el_quickinfor_testid.text.strip()!=testid:
                showinfo(generate_report.report_status,'清空输入框历史输入记录...',"black")
                el_input = driver.find_element(By.XPATH,laboratories_input)
                el_input.clear()
                time.sleep(1)
                showinfo(generate_report.report_status,'输入测试单号...',"black")
                el_input.send_keys(self.testid_input.get())
                showinfo(generate_report.report_status,'提交测试单号...',"black")
                el_input.send_keys('\n')
                time.sleep(2)
                testid1=self.testid_input.get().replace('.','_')
                testid_item = (By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]')
                wait_element_text(testid_item,'testid_item',testid1,15)
                process=(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div') #进程显示图标   
                #等待进程显示图标消失
                wait_element_invisibility(process,'testid_process',30)
                showinfo(generate_report.report_status,'点选查找出来的测试项...',"black")
                el_testid = driver.find_element(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]').click()  #模糊查找   
                quickinfor_testid=(By.XPATH,laboratories_quickinfo)
                wait_element_text((By.XPATH,laboratories_quickinfo),'quickinfo',testid,15)  
                time.sleep(4)  
            #点击result_files    
            showinfo(generate_report.report_status,'点击result_files',"black")
            driver.find_element(By.XPATH,laboratories_result_files).click()
            #点击new
            wait_element_clickable((By.XPATH,laboratories_new),'new',15)
            showinfo(generate_report.report_status,'点击new',"black")
            time.sleep(2)
            # el_new=driver.find_element(By.XPATH,laboratories_new)
            # #必须将鼠标移动至new并等待三秒,否则跳出的new提示框会导致点击result file报错
            # ActionChains(driver).move_to_element(el_new).perform()
            # time.sleep(3)
            # el_new.click()
            driver.find_element(By.XPATH,laboratories_new).click()
            #点击result file
            showinfo(generate_report.report_status,'点击result file',"black")
            time.sleep(1)
            driver.find_element(By.XPATH,laboratories_result_file_img).click()
            time.sleep(1)
            #输入title
            showinfo(generate_report.report_status,'输入result file title',"black")
            title_input=driver.find_element(By.XPATH,laboratories_title)
            title_input.clear()
            time.sleep(1)
            title_input.send_keys(title)
            time.sleep(1)
            if needreport==0:
                showinfo(generate_report.report_status,'输入comment',"black")
                comment_input=driver.find_element(By.XPATH,laboratories_comment)
                comment_input.clear()
                time.sleep(1)
                comment_input.send_keys('no need report')
                time.sleep(1)
            #提交link地址
            showinfo(generate_report.report_status,'提交上传文件地址',"black")
            driver.find_element(By.XPATH,laboratories_add_document).send_keys(file_dir)
            time.sleep(1)
            #点击save保存
            showinfo(generate_report.report_status,'点击save',"black")
            driver.find_element(By.XPATH,laboratories_save).click()
            time.sleep(2)
            showinfo(generate_report.report_status,f'{testid} 文件上传完毕',"green")
            tkinter.messagebox.showinfo('提示', f'{testid} 文件上传完毕')
        except Exception as e:
            showinfo(generate_report.report_status,'上传文件出现异常',"red")
            tkinter.messagebox.showinfo('提示', f'上传文件出现异常,请重新点击上传')
            max_progress_value=0
            sys.exit(1)


    #创建表格        
    def create_table(self,test_data_path):
        try:
            file_list = os.listdir(test_data_path)
            heading_name = ''
            column_number = 0  #表格列数
            row_number = 0     #表格行数
            excess_cell_number=0 #多余的表格单元格数
            data_type = ''
            for file in file_list:
                filter_data_list = []
                if os.path.isdir(test_data_path + '/' + file):  # 如果文件为文件夹
                    data_type = file.split('_')[0][0]
                    if data_type == '1':
                        heading_name = 'Picture(s) about the test setup'
                    elif data_type == '2':
                        heading_name = 'Test process records'
                    elif data_type == '3':
                        heading_name = 'Test results and evaluation'
                    elif data_type == '4':
                        break
                    data_list = os.listdir(test_data_path + '\\' + file)
                    for data in data_list:  # 过滤隐藏文件和文件夹
                        data_path = test_data_path + '\\' + file + '\\' + data
                        if not file_is_hidden(data_path) and not os.path.isdir(data_path) and is_file_image(data_path):
                            filter_data_list.append(data)
                        if is_excel_file(data) and data_type!='4':
                            for o in range(len(doc.paragraphs)):
                                if (doc.paragraphs[o]).text == heading_name:
                                    doc.paragraphs[o+1].insert_paragraph_before('') 
                                    doc.paragraphs[o+ 1].insert_paragraph_before(data.split('.')[0])  # 在某段插入excel文件名

                    data_number = len(filter_data_list)  # 计算数据数量
                    column_number = int(file.split('_')[0][2])
                    if data_number == 0:  # 如果数量为0则说明此文件夹无数据 不需要添加
                        continue
                    if data_number < column_number:
                        row_number = 3
                        column_number = data_number
                        excess_cell_number=column_number-data_number
                    elif data_number % column_number == 0:
                        row_number = ((data_number // column_number) * 2) + 1
                        excess_cell_number=0
                    else:
                        row_number = ((data_number // column_number + 1) * 2) + 1
                        excess_cell_number=column_number-data_number%column_number

                    # 在指定位置创建表格
                    for n in range(len(doc.paragraphs)):
                        if (doc.paragraphs[n]).text == heading_name:
                            doc.paragraphs[n + 1].insert_paragraph_before('')  # 在表格之前插入一个空白行
                            table = doc.add_table(row_number, column_number, style="Table Grid")
                            #delete_cell_border(table,column_number,row_number,excess_cell_number) #此处不能删除多余单元格否则后面插入图片会报错
                            table.alignment = WD_TABLE_ALIGNMENT.CENTER
                            # 表格首行合并单元格
                            for m in range(column_number):
                                table.cell(0, 0).merge(table.cell(0, m))
                            # 如果只有一列则图片居中显示
                            if column_number == 1:
                                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                            # 生成的表格宽度自适应图片大小
                            for row in table.rows:
                                for cell in row.cells:
                                    cell.width = Inches(0)
                            # 设置表格第一行文本显示以及样式
                            run = table.cell(0, 0).paragraphs[0].add_run(file.split('_')[1])
                            run.bold = True
                            run.italic = True
                            doc.paragraphs[n + 1]._element.addnext(table._tbl)
        except Exception as e:
            showinfo(generate_report.report_status,f'表格创建异常:{str(e)}',"red")
            tkinter.messagebox.showinfo('提示', '表格创建异常,程序终止')
            max_progress_value=0
            sys.exit(1)

    #插入数据
    def insert_data(self,test_data_path):
        #创建容纳图片的表格
        self.create_table(test_data_path)

        data_folder_list = os.listdir(test_data_path)
        # data_folder_list.sort(reverse=False)
        table_list = doc.tables
        for data_folder in data_folder_list:
            if os.path.isdir(test_data_path + '\\' + data_folder):
                data_type = data_folder.split('_')[0][0]           
                if data_type == 4:
                    break
                data_folder_name = data_folder.split('_')[1]
                column_number = int(data_folder.split('_')[0][2])
                data_list = os.listdir(test_data_path + '\\' + data_folder)
                filter_data_list = []
                for file in data_list:
                    file_path = test_data_path + '\\' + data_folder + '\\' + file
                    if not file_is_hidden(file_path) and not os.path.isdir(file_path) and is_file_image(file_path):
                        filter_data_list.append(file)
                if len(filter_data_list) == 0:
                    showinfo(generate_report.report_status,f'{data_folder}无需插入数据',"black")      
                else: 
                    if self.time_sort.get()==1: #按数据创建时间排序
                        filter_data_list.sort(key=lambda x:os.path.getmtime(test_data_path + '\\' + data_folder + '\\' + x))
                    data_number = len(filter_data_list)
                    for table in table_list:
                        if table.cell(0, 0).text == data_folder_name:
                            row_number=0
                            column_number=0
                            excess_cell_number=0   #多余单元格
                            row_number=len(table.rows) #行数
                            column_number=len(table.columns)  #列数
                            if data_number < column_number:
                                excess_cell_number=column_number-data_number
                            elif data_number % column_number == 0:
                                excess_cell_number=0
                            else:
                                excess_cell_number=column_number-data_number%column_number
                            for j, data in enumerate(filter_data_list):
                                try:
                                    cell = table.cell(((j // column_number)) * 2 + 1, j % column_number)
                                    paragraph = cell.paragraphs[0]
                                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                                    run = paragraph.add_run()
                                    pic = Image.open(test_data_path + '\\' + data_folder + '\\' + data)  # 载入图片文件
                                    c = pic.width / pic.height  # 图片的宽度和高度比例
                                    if len(filter_data_list) == 1:
                                        pic_width = 16
                                        pic_height = pic_width / c
                                    else:
                                        pic_width = 9.3 * (2 / column_number)
                                        pic_height = pic_width / c
                                    pic_size = (os.path.getsize(
                                        test_data_path + '\\' + data_folder + '\\' + data)) // 1024
                                    if pic_size > 1000 and self.compress_picture.get() == 1:
                                        compress_picture_path = self.compress_picture_size(
                                            test_data_path + '\\' + data_folder + '\\' + data, 50)
                                        run.add_picture(compress_picture_path, width=Cm(pic_width), height=Cm(pic_height))
                                    else:
                                        run.add_picture(test_data_path + '\\' + data_folder + '\\' + data,
                                                        width=Cm(pic_width), height=Cm(pic_height))
                                    data_name = data.rsplit('.',1)[0]
                                    paragraph = table.cell(((j // column_number)) * 2 + 2, j % column_number).paragraphs[0]
                                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                                    paragraph.add_run(data_name)
                                except Exception as e:
                                    showinfo(generate_report.report_status,f'{data_folder}里的{data}图片插入失败:{str(e)}',"red")
                                else:
                                    showinfo(generate_report.report_status,f'{data_folder}里的{data}图片插入成功',"black")
                            #删除多余单元格    
                            self.delete_cell_border(table,column_number,row_number,excess_cell_number)       
                            break
        
    #检查指定数据文件夹内是否存在附件,如果存在则返回包含所有附件的数组
    def check_attachment(self,test_data_path):
        attachment_list=[]
        for file in os.listdir(test_data_path):
            if file.split('_')[0][0] == '4': #附件
                for data in os.listdir(test_data_path + '\\' + file):
                    if not file_is_hidden(test_data_path + '\\' + file+'\\'+data) and not os.path.isdir(test_data_path + '\\' + file+'\\'+data):
                        attachment_list.append(test_data_path + '\\' + file+'\\'+data)
        return attachment_list
            
    #检查指定数据文件夹内是否存在需要插入excel文件数据,如果存在则将excel文件所在文件夹名称的第1，2位数字以及excel文件的地址存入excel_file_list
    def check_excel(self,test_data_path):
        excel_file_list=[]
        for file in os.listdir(test_data_path):
            file_path=test_data_path + '\\' + file
            if os.path.isdir(file_path) and file.split('_')[0][0]!='4':
                    datalist=os.listdir(file_path)
                    for data in datalist:
                        data_path=test_data_path + '\\' + file+'\\'+data
                        if is_excel_file(data_path) and not file_is_hidden(data_path):
                            first=file.split('_')[0][0]  #文件夹名称第一个数字:代表插入位置
                            second=file.split('_')[0][1] #文件夹名称第二个数字:代表插入顺序
                            excel_file_list.append((first,second,data_path))
        return excel_file_list                    

    #插入结论和审批人信息                        
    def add_conclusion_approval(self):
        # if V1.get() == 0:
        #     maintest_result = 'positive'
        # else:
        #     maintest_result = 'negative'
        try:
            paragraphs = doc.paragraphs
            paragraph = paragraphs[2]
            paragraph.text = ''
            paragraphs[3].text = ''
            paragraphs[4].text = ''
            # paragraphs[1].text = 'After vibration test:'
            # paragraphs[1].text = combobox34.get()
            conclusions = self.pre_conclusion_select.current_value
            conclusions_list = conclusions.split('&')
            for conclusion in conclusions_list:
                run = paragraph.add_run(conclusion + '\n')
                run.font.size = Pt(12)
            # paragraphs[1].runs[0].font.size = Pt(12)

            if len(StringVarlist1) != 0:
                run = paragraphs[2].add_run('post test:')
                run.font.size = Pt(12)
                run.bold = True
            for i in range(len(StringVarlist1)):
                if IntVarlist[i].get() == 0:
                    posttest_result = 'positive'
                else:
                    posttest_result = 'negative'
                text = 'The ' + StringVarlist2[i].get() + ' is ' + posttest_result + ',detail test data refer to the ' + \
                    StringVarlist1[i].get()
                paragraphs[3].insert_paragraph_before(text).runs[0].font.size = Pt(12)
            approval_table = doc.sections[0].first_page_footer.tables[0]
            reviewer = self.first_approval.get()
            approval = self.second_approval.get()
            approval_table.cell(2, 1).text = reviewer
            approval_table.cell(3, 1).text = approval
            
            # #在Test results and evaluation处增加结论描述
            # conclusions = post_conclusion_select.current_value
            # conclusions_list = conclusions.split('&')
            # for i in range(len(paragraphs)):
            #     if paragraphs[i].text=="Test results and evaluation":
            #         p=paragraphs[i+1].insert_paragraph_before()
            #         for conclusion in conclusions_list:
            #             run = p.add_run(conclusion + '\n')
            #             run.font.size = Pt(12)

        except Exception as e:
            showinfo(generate_report.report_status,"插入结论和审批人信息失败","red")
            showinfo(generate_report.report_status,str(e),"red")

    #插入设备信息
    def insert_equipment(self):
        try:
            table = doc.tables[0]
            paragraph = table.cell(14, 1).paragraphs[0]
            # p=paragraph._element
            # p.getparent().remove(p)
            # paragraph._p=paragraph._element=None
            paragraph.clear()
            equipments = self.equipment_select.current_value
            equipments_list = equipments.split('&')
            current_date=datetime.now()
            for equipment in equipments_list:
                calibration_date=equipment.split('/',1)[1].split('(')[1].split(')')[0]  #计量有效期
                comparison_date=datetime.strptime(calibration_date,'%Y-%m-%d')
                if comparison_date <= current_date:             #如果当前日期超过计量日期，则插入的设备信息文字标红
                    new_paragraph=paragraph.insert_paragraph_before()
                    run=new_paragraph.add_run(equipment)
                    run.font.color.rgb=RGBColor(255,0,0)
                else:    
                    paragraph.insert_paragraph_before(equipment)
            if self.lab_temperature.get()!=''and self.lab_humidity.get()!='':
                paragraph.insert_paragraph_before(f"lab temerature and humidity:{self.lab_temperature.get()}°C,{self.lab_humidity.get()}%RH")
        except Exception as e:
            showinfo(generate_report.report_status,"插入设备信息失败","red")
            showinfo(generate_report.report_status,str(e),"red")

    #插入文档数据
    def insert_document(self,test_data_path):
        global excel_file_list
        win_word=None
        win_excel=None
        word_doc=None
        #确认是否需要插入附件
        showinfo(generate_report.report_status,'检查是否需要插入附件....',"black")
        attachment_list=self.check_attachment(test_data_path)
        if len(attachment_list) !=0: #存在附件
            doc.add_paragraph(' ')
            delete_content("The attachments can be found from the next page.")
            doc.save(report_dir)    
            pythoncom.CoInitialize()
            try:
                showinfo(generate_report.report_status,'需要插入附件,打开word进程中....',"black")
                win_word = win32com.client.DispatchEx("Word.Application")
                win_word.DisplayAlerts=False
                win_word.Visible = False
                word_doc = win_word.Documents.Open(report_dir)
                showinfo(generate_report.report_status,'word进程已经打开,插入附件中...',"black")
                self.add_attachment(attachment_list,word_doc)
                showinfo(generate_report.report_status,'附件插入完毕',"green")
            except Exception as e:
                showinfo(generate_report.report_status,f'打开word进程或者word文档的过程中发生异常:{str(e)}',"red")
            finally: 
                try:
                    showinfo(generate_report.report_status,'保存和关闭word文档中...',"black")
                    if word_doc:
                        word_doc.Save()
                        word_doc.Close()
                except Exception as e:
                    showinfo(generate_report.report_status,f'保存和关闭word文档时异常:{str(e)}',"red")
                    pass
                try:
                    showinfo(generate_report.report_status,'关闭word进程中...',"black")
                    win_word.Quit()
                except Exception as e:
                    showinfo(generate_report.report_status,f'关闭word进程时异常:{str(e)}',"red")
                    pass    
                pythoncom.CoUninitialize()   
        else:  #不存在附件
            delete_content("The attachments can be found from the next page.")
            doc.add_paragraph().add_run('There are no attachments.')
            doc.save(report_dir)                 
        #确认是否需要插入excel数据
        showinfo(generate_report.report_status,'检查是否需要插入excel数据....',"black")
        excel_file_list=self.check_excel(test_data_path)
        if len(excel_file_list) !=0: #存在excel数据
            showinfo(generate_report.report_status,'需要插入excel数据',"black")
            pythoncom.CoInitialize()
            try:
                showinfo(generate_report.report_status,'打开word进程中....',"black")
                win_word = win32com.client.DispatchEx("Word.Application")
                win_word.DisplayAlerts=False
                win_word.Visible = False
                word_doc = win_word.Documents.Open(report_dir)
                showinfo(generate_report.report_status,'打开excel进程中....',"black")
                win_excel = win32com.client.DispatchEx("Excel.Application")  #此处使用DispatchEx额外打开一个进程这样在关闭进程时不会将其它excel文件关闭  
                win_excel.DisplayAlerts=False  
                win_excel.Visible = False  #让excel关闭时不跳出对话提示框
                showinfo(generate_report.report_status,'插入excel数据中....',"black")
                self.add_excel_data(excel_file_list,win_excel,word_doc)   
                showinfo(generate_report.report_status,'excel数据插入完毕',"green") 
            except Exception as e:
                showinfo(generate_report.report_status,f'插入excel数据的过程中发生异常:{str(e)}',"red")
                print(str(e))    
            finally:
                try:
                    showinfo(generate_report.report_status,'保存和关闭word文档中...',"black")
                    if word_doc:
                        word_doc.Save()
                        word_doc.Close()
                except Exception as e:
                    showinfo(generate_report.report_status,f'保存和关闭word文档时异常:{str(e)}',"red")
                    pass
                try:
                    showinfo(generate_report.report_status,'关闭word进程中...',"black")
                    win_word.Quit()
                except Exception as e:
                    showinfo(generate_report.report_status,f'关闭word进程时异常:{str(e)}',"red")
                    pass    
                try:
                    showinfo(generate_report.report_status,'关闭excel进程中...',"black")
                    win_excel.Quit()
                except Exception as e:
                    showinfo(generate_report.report_status,f'关闭excel进程时异常:{str(e)}',"red")
                    pass   
                pythoncom.CoUninitialize() 

    #插入附件
    def add_attachment(self,attachment_list,word_doc):   
        paragraphlist = word_doc.Paragraphs
        total = len(paragraphlist)
        ico_file_path1=os.path.join(current_dir,"PDFFile_8.ico")  #pdf文件图标
        ico_file_path2=os.path.join(current_dir,"Excel.ico")      #excel文件图标
        ico_file_path3=os.path.join(current_dir,"txt.ico")        #text文件图标
        ico_file_path4=os.path.join(current_dir,"word.ico")       #word文件图标
        ico_file_path5=os.path.join(current_dir,"pic.ico")        #图片文件图标
        ico_file_path6=os.path.join(current_dir,"email.ico")      #邮件文件图标
        ico_file_path=''
        for attachment in attachment_list:
            attachment_name=os.path.basename(attachment).split('.')[0]
            attachment_suffix=os.path.basename(attachment).split('.')[1]
            try:
                if attachment_suffix in ['xlsx','xls']:
                    ico_file_path=ico_file_path2
                if attachment_suffix in['pdf']:
                    ico_file_path=ico_file_path1
                if attachment_suffix in['txt']:
                    ico_file_path=ico_file_path3    
                if attachment_suffix in['docx','doc']:
                    ico_file_path=ico_file_path4      
                if attachment_suffix in['jpg','jpeg','png','bmp']:    
                    ico_file_path=ico_file_path5
                if attachment_suffix in['msg']:    
                    ico_file_path=ico_file_path5
                attachment_data = paragraphlist(total).Range.InlineShapes.AddOLEObject(
                    ClassType="AcroExch.Document",
                    FileName=attachment,
                    LinkToFile=False,
                    DisplayAsIcon=True,
                    IconFileName=ico_file_path,
                    IconIndex=13,
                    IconLabel=attachment_name
                    )
            except Exception as e:
                showinfo(generate_report.report_status,f'{attachment_name}附件插入失败:{str(e)}',"red")
            else:
                showinfo(generate_report.report_status,f'{attachment_name}附件插入成功',"green")

    #插入excel数据        
    def add_excel_data(self,excel_file_list,win_excel,word_doc):
        #首先将各个excel文件名写入attachment处进行占位
        for excel_file in excel_file_list:
            excel_file_name=os.path.basename(excel_file[2]).split('.')[0]
            for p in word_doc.Paragraphs:
                if p.Range.Text.strip()=='Attachments':
                    target_range=p.Range
                    target_range.InsertAfter('\n'+excel_file_name)
                    target_range.InsertAfter('\n')      ##此处必须插入一个空行 否则excel数据插入失败
        for excel_file in excel_file_list:
            try:
                excel_file_name=os.path.basename(excel_file[2]).split('.')[0]
                workbook=win_excel.Workbooks.Open(excel_file[2])
                worksheet=workbook.Worksheets(1)          #此处打开第一个sheet
                copy_area=worksheet.Cells(1,1).Value 
                range_to_copy = worksheet.Range(copy_area)
                #range_to_copy.CopyPicture()
                range_to_copy.Copy()
                for i,p in enumerate(word_doc.Paragraphs):
                    if p.Range.Text.strip()==excel_file_name.strip():
                        p.Range.Paste()
                        time.sleep(1)
                    #此处必须打开剪切板并将剪切板中内容清空，否则会跳出对话提示框
                        win32clipboard.OpenClipboard()
                        win32clipboard.EmptyClipboard()
                        win32clipboard.CloseClipboard()
            except Exception as e:
                showinfo(generate_report.report_status,f'{excel_file_name}excel数据插入失败:{str(e)}',"red")
            else:
                showinfo(generate_report.report_status,f'{excel_file_name}数据插入成功',"green")
            finally:
                try:
                    workbook.Close(SaveChanges=False)   
                except Exception as e:     
                    showinfo(generate_report.report_status,f'关闭excel文档时异常:{str(e)}',"red")
                    pass

    #增加后测试    
    def add_post_test(self):
        i = len(self.posttest_frame.winfo_children())
        S1 = tkinter.StringVar()
        S2 = tkinter.StringVar()
        entry2 = tkinter.Entry(self.posttest_frame, width=12, textvariable=S1).grid(row=(i // 3) + 3, column=1)
        StringVarlist1.append(S1)
        entry3 = tkinter.Entry(self.posttest_frame, width=12, textvariable=S2).grid(row=(i // 3) + 3, column=2, sticky=W)
        StringVarlist2.append(S2)
        V = tkinter.IntVar()
        radiobutton1 = tkinter.Radiobutton(self.posttest_frame, text='OK', width=3, variable=V, value=0).grid(row=(i // 3) + 3, column=3, sticky=W)
        radiobutton2 = tkinter.Radiobutton(self.posttest_frame, text='NOK', width=3, variable=V, value=1).grid(row=(i // 3) + 3, column=4, sticky=W)
        IntVarlist.append(V)

    #删除后测试
    def delete_post_test(self):
        childrenlist = self.posttest_frame.winfo_children()
        i = len(childrenlist)
        if childrenlist[i - 1].winfo_class() == 'Entry' or childrenlist[i - 1].winfo_class() == 'Radiobutton':
            childrenlist[i - 1].destroy()
        if childrenlist[i - 2].winfo_class() == 'Entry' or childrenlist[i - 2].winfo_class() == 'Radiobutton':
            childrenlist[i - 2].destroy()
            IntVarlist.pop()
        if childrenlist[i - 3].winfo_class() == 'Entry' or childrenlist[i - 3].winfo_class() == 'Radiobutton':
            childrenlist[i - 3].destroy()
            StringVarlist2.pop()
        if childrenlist[i - 4].winfo_class() == 'Entry' or childrenlist[i - 4].winfo_class() == 'Radiobutton':
            childrenlist[i - 4].destroy()
            StringVarlist1.pop()

        #更新测试规范
    def update_test_spec(self):
        test_id = self.testid_input.get()
        if test_id=='':
            tkinter.messagebox.showinfo('提示', '输入空值无效')
            sys.exit(1)
        test_folder_path = search_folder3(test_dir, test_id)
        if test_folder_path is None:
            showinfo(generate_report.report_status,'测试文件夹不存在',"red")
            tkinter.messagebox.showinfo('提示', '测试文件夹不存在')
            sys.exit(1)
        else:
            download_test_spec(test_id)
            download_file = get_download_file(15)
            if download_file is not None:
                showinfo(generate_report.report_status,'测试规范下载完毕,移动测试规范中...',"black")
                move_file(download_file, test_folder_path,1)    
                showinfo(generate_report.report_status,f'{test_id}测试规范更新完毕',"green")
            else:
                showinfo(generate_report.report_status,f'超时,{test_id}测试规范下载失败,请检查网络或其他原因',"red")
                tkinter.messagebox.showinfo('提示', '测试规范下载失败,请检查网络或其他原因')
                sys.exit(1)     
      
    def get_test_detail(self):
        edateapi=EdateApi()
        testid = self.testid_input.get()
        entityid=''
        search_key=update_test_plan.search_test_criteria1.get()
        resultlist1=edateapi.get_test_orders_by_filter(search_key,'ZHS6CNG')
        for result in resultlist1:
            if result['ID']==testid:
                entityid=result['id']    
        if entityid !='':        
            resultlist2=edateapi.get_test_order_details(entityid)
            return resultlist2[0]
        else:
            return None    

        #更新测试状态        
    def update_test_status(self):
        global driver
        global edge_options
        global max_progress_value
        global visit_times              #访问页面次数
        visit_times=visit_times+1
        testid = self.testid_input.get()
        test_status=self.test_status.get()
        if testid=='' or test_status=='':
            tkinter.messagebox.showinfo('提示', f'请输入一个testid和选择一个状态') 
            return
        showinfo(generate_report.report_status,'正在获取输入测试单号的状态信息,请稍等...',"green")
        result=self.get_test_detail()
        if result!=None:
            current_test_status=result['TestInstance#status']
            planned_start=result['TestInstance#plannedStart']
            planned_end=result['TestInstance#plannedEnd']
            current_object_status=result['TestObject#status']
            current_object_deliverdate=result['TestObject#deliveryDate']
            Orderer=result['Orderer#displayName']    
            Responsible=result['Responsible#displayName'] 
            showinfo(generate_report.report_status,f'current test status:{current_test_status}',"green")
            showinfo(generate_report.report_status,f'planned_start:{planned_start}',"green")
            showinfo(generate_report.report_status,f'planned_end:{planned_end}',"green")
            showinfo(generate_report.report_status,f'current object status:{current_object_status}',"green")
            showinfo(generate_report.report_status,f'sample deliver date:{current_object_deliverdate}',"green")
            showinfo(generate_report.report_status,f'Order:{Orderer}',"green")
            showinfo(generate_report.report_status,f'Responsible:{Responsible}',"green")

            if test_status=='Query only':
                return
            if test_status=='PERFORMED' and current_test_status in ('INQUIRED','PLANNED'):
                showinfo(generate_report.report_status,'请输入试验开始和结束时间...',"red")
                data_range_picker=DateRangePicker(window)
                result,start_date,end_date=data_range_picker.show(current_object_deliverdate)
                if result==False:
                    tkinter.messagebox.showinfo('提示', f'必须选择日期否则无法完成修改')  
                    sys.exit(1)

            if testid=='':
                tkinter.messagebox.showinfo('提示', '输入空值无效')    
            #['CREATED','INQUIRED','PLANNED','TEST_OBJECTS_AVAILABLE','IN_PROGRESS','PERFORMED','ASSESSED']
            elif test_status==current_test_status:
                tkinter.messagebox.showinfo('提示', f'当前状态已经是{current_test_status},无需更改')
            elif test_status=='CREATED' and current_test_status in ('IN_PROGRESS','PERFORMED','ASSESSED'):
                tkinter.messagebox.showinfo('提示', f'当前状态是{current_test_status},无法直接更改成CREATED状态,请先更改至前一个状态')   
            elif test_status=='INQUIRED' and current_test_status in ('TEST_OBJECTS_AVAILABLE','IN_PROGRESS','PERFORMED','ASSESSED','ASSESSED_REVIEWED'):
                tkinter.messagebox.showinfo('提示', f'当前状态是{current_test_status},无法直接更改成INQUIRED状态,请先更改至前一个状态')
            elif test_status=='PLANNED' and current_test_status in ('IN_PROGRESS','PERFORMED','ASSESSED','ASSESSED_REVIEWED'):   
                tkinter.messagebox.showinfo('提示', f'当前状态是{current_test_status},无法直接更改成PLANNED状态,请先更改至前一个状态')
            elif test_status=='PLANNED' and current_test_status=='INQUIRED' and planned_start=='':
                tkinter.messagebox.showinfo('提示', f'需要排具体测试时间,请手动至e-data更改')   
            elif test_status=='TEST_OBJECTS_AVAILABLE'and current_test_status=='PLANNED' and current_object_deliverdate=='':
                tkinter.messagebox.showinfo('提示', f'试验委托人还未确认样品delivery,无法更改状态至TEST_OBJECTS_AVAILABLE')   
            elif test_status=='TEST_OBJECTS_AVAILABLE'and current_test_status=='INQUIRED':        
                tkinter.messagebox.showinfo('提示', f'当前状态是INQUIRED,请先改成PLANNED')
            elif test_status=='TEST_OBJECTS_AVAILABLE'and current_test_status in ('PERFORMED','ASSESSED','ASSESSED_REVIEWED'):    
                tkinter.messagebox.showinfo('提示', f'当前状态是{current_test_status},无法直接更改成TEST_OBJECTS_AVAILABLE状态,请先更改至前一个状态')
            elif test_status=='IN_PROGRESS' and current_test_status in ('ASSESSED','ASSESSED_REVIEWED'):
                tkinter.messagebox.showinfo('提示', f'当前状态是{current_test_status},无法直接更改成IN_PROGRESS状态,请先更改至前一个状态')
            elif test_status=='IN_PROGRESS' and current_object_status!='DELIVERED':
                tkinter.messagebox.showinfo('提示', f'试验委托人还未确认样品delivery,无法更改状态至IN_PROGRESS')
            elif test_status=='PERFORMED' and current_test_status=='TEST_OBJECTS_AVAILABLE':
                tkinter.messagebox.showinfo('提示', f'当前状态是TEST_OBJECTS_AVAILABLE,需要先更改至IN_PROGRESS状态')    
            else:
                if isElementExist(alert_button):  #如果存在页面提示框则点击提示框ok
                    driver.find_element(By.XPATH,alert_button).click()
                    time.sleep(2)
                if driver==None or driver.current_url.split(';')[0]!=laboratories.split(';')[0]:
                    initial_browser(laboratories)
                    wait_element_clickable((By.XPATH, laboratories_open),'open',30)
                    time.sleep(7) #必须
                try:
                    showinfo(generate_report.report_status,'定位测试单号输入框...',"black")
                    el_input = driver.find_element(By.XPATH,laboratories_input)
                    # 获取输入框标签
                    if el_input.get_attribute("value")!=testid or visit_times==1:
                        showinfo(generate_report.report_status,'清空输入框历史输入记录...',"black")
                        el_input.clear()
                        showinfo(generate_report.report_status,'输入测试单号...',"black")
                        el_input.send_keys(testid)
                        showinfo(generate_report.report_status,'提交测试单号...',"black")
                        el_input.send_keys('\n')
                        time.sleep(2)
                        testid1=testid.replace('.','_') 
                        testid_item = (By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]')
                        wait_element_text(testid_item,'testid_item',testid1,15)
                        process=(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div') #进程显示图标
                        #等待进程显示图标消失
                        wait_element_invisibility(process,'testid process',30)
                        time.sleep(2)
                        showinfo(generate_report.report_status,'点选查找出来的测试项...',"black")
                        el_testid = driver.find_element(By.XPATH,'/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]//span[contains(text(),"' + testid1 + '")]').click()  #模糊查找
                        time.sleep(1)                                     
                        showinfo(generate_report.report_status,'等待Quick Infor页面元素刷新显示完毕...',"black")
                        quickinfor_testid=(By.XPATH,laboratories_quickinfo)
                        wait_element_text((By.XPATH,laboratories_quickinfo),'quickinfo',testid,15) 
                        time.sleep(1)       
                    # #获取test_status
                    # el_laboratories_qucikinfo_status=driver.find_element(By.XPATH,laboratories_qucikinfo_test_status)  
                    # current_test_status=el_laboratories_qucikinfo_status.text.strip()       
                    # time.sleep(1)  
                    # #获取object_status和object_delivery_date
                    # el_laboratories_overview_object=driver.find_element(By.XPATH,laboratories_overview_object)
                    # el_laboratories_overview_object.click()  
                    # time.sleep(2)
                    # el_laboratories_qucikinfo_object_status=driver.find_element(By.XPATH,laboratories_qucikinfo_object_status)
                    # current_object_status=el_laboratories_qucikinfo_object_status.text.strip()
                    # el_laboratories_qucikinfo_object_deliverdate=driver.find_element(By.XPATH,laboratories_qucikinfo_object_deliverdate)
                    # current_object_deliverdate=el_laboratories_qucikinfo_object_deliverdate.text.strip()  
                    # el_laboratories_overview_object_ico=driver.find_element(By.XPATH,laboratories_overview_object_ico)
                    # el_laboratories_overview_object_ico.click()                                   #点击第一行样品信息最左边选择框取消被选择状态从而回到使得quick info回到test info
                    # time.sleep(2)
                    driver.find_element(By.XPATH,laboratories_navigation_testitem).click()    #此处必须点击一下 确保改变状态的三角形符号出现
                    time.sleep(2)
                    driver.find_element(By.XPATH,laboratories_navigation_ico).click()      #点击状态更改标签的三角形符号
                    time.sleep(2)
                    driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus).click()   #点击change status图标
                    time.sleep(2)

                    if test_status=='CREATED' and current_test_status in ('INQUIRED','PLANNED','TEST_OBJECTS_AVAILABLE'):
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_created).click()
                        time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_created_created).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至CREATED')
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")

                    elif test_status=='INQUIRED' and current_test_status=='PLANNED':
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_inquired).click()
                        time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_inquired_inquired).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至INQUIRED')
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")

                    elif test_status=='PLANNED' and current_test_status=='TEST_OBJECTS_AVAILABLE':             
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_planned).click()
                        time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_planned_planned).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至PLANNED')
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")

                    elif test_status=='PLANNED' and current_test_status=='INQUIRED' and planned_start!='':     
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_planned).click()
                        time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_planned_planned).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至PLANNED')
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")

                    elif test_status=='TEST_OBJECTS_AVAILABLE' and current_test_status=='PLANNED' and current_object_deliverdate!=''  :
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_testobjectsavailable).click()
                        time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_testobjectsavailable_ok).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至TEST_OBJECTS_AVAILABLE')
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")

                    elif test_status=='TEST_OBJECTS_AVAILABLE' and current_test_status=='IN_PROGRESS':
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_testobjectsavailable).click()
                        time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_testobjectsavailable_ok).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至TEST_OBJECTS_AVAILABLE')
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")

                    elif test_status=='IN_PROGRESS' and current_test_status in ('INQUIRED','PLANNED','TEST_OBJECTS_AVAILABLE') and current_object_status=='DELIVERED':
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_inprogress).click()
                        time.sleep(2)
                        # current_date = datetime.now().strftime("%Y-%m-%d")
                        # print(current_date)
                        # driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_inprogress_startdate).send_keys(current_date)
                        # time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_inprogress_start).click()
                        time.sleep(2) 
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至IN_PROGRESS')  
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")

                    elif test_status=='IN_PROGRESS' and current_test_status=='PERFORMED':
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_inprogress).click()
                        time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_inprogress_start).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至IN_PROGRESS')
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")

                    elif test_status=='PERFORMED' and current_test_status in ('INQUIRED','PLANNED'):
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_performed).click()
                        time.sleep(2)
                        start_date_input=driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_performed_startdate)
                        start_date_input.clear()
                        time.sleep(1)
                        start_date_input.send_keys(start_date)
                        time.sleep(2)
                        end_date_input=driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_performed_enddate)
                        end_date_input.clear()
                        time.sleep(1)
                        end_date_input.send_keys(end_date)
                        time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_performed_complete).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至PERFORMED')  
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")
                    elif test_status=='PERFORMED' and current_test_status=='IN_PROGRESS':
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_performed).click()
                        time.sleep(2)
                        # driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_performed_startdate).send_keys(end_date)
                        # time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_performed_complete).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至PERFORMED')  
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")
                    elif test_status=='PERFORMED' and current_test_status=='ASSESSED':    
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_performed).click()
                        time.sleep(2)
                        driver.find_element(By.XPATH,laboratories_navigation_ico_changestatus_performed_complete).click()
                        time.sleep(2)
                        tkinter.messagebox.showinfo('提示', f'已修改{testid}状态至PERFORMED')  
                        showinfo(generate_report.report_status,'修改测试状态完毕',"green")
                    else:
                        tkinter.messagebox.showinfo('提示', '暂不支持该操作')

                except Exception as e:
                    tkinter.messagebox.showinfo('提示', '页面操作过程中报错，请点击重试')
                    print(str(e))
        else:
            showinfo(generate_report.report_status,'未查询到该单号,请检查搜索条件...',"red")
            tkinter.messagebox.showinfo('提示', '未查询到该单号,请检查搜索条件...')               


class UpdateTestPlan:
    def __init__(self, root):
        self.root = root
        # 变量
        self.cobo_test_type=None                                                #试验类型
        self.cobo_register_type=None                                            #登记类型
        self.update_testinfo = tkinter.IntVar()                                 #更新测试信息表勾选值
        self.update_testinfo.set(update_testinfo_current_value)
        self.create_folder = tkinter.IntVar()                                   #创建测试文件夹勾选值
        self.create_folder.set(create_folder_current_value)
        self.if_download_spec = tkinter.IntVar()                                   #下载测试规范勾选值
        self.if_download_spec.set(download_spec_current_value)
        self.cb_download_spec=None
        

        self.with_foldername = tkinter.IntVar()                                 #文件夹附带名称勾选值
        self.with_foldername.set(with_foldername_current_value)
        self.with_foldername1 = tkinter.IntVar()                                 #文件夹附带名称勾选值
        self.with_foldername1.set(with_foldername1_current_value)
        self.cb_with_foldername=None
        self.search_test_criteria = tkinter.StringVar()                         #测试信息搜索条件名
        self.search_test_criteria.set(search_test_criteria_current_value)
        self.search_object_criteria = tkinter.StringVar()                       #样品信息搜索条件名
        self.search_object_criteria.set(search_object_criteria_current_value)

        self.search_test_criteria1 = tkinter.StringVar()                        #测试信息搜索条件名2
        self.search_test_criteria1.set(search_test_criteria1_current_value)

        self.sheet_name = tkinter.StringVar()                                   #计划表分页
        self.sheet_name.set(sheet_name_current_value)              
        self.begin_row = tkinter.StringVar()                                    #计划表起始行
        self.begin_row.set(begin_row_current_value)
        self.hyperlink_column = tkinter.StringVar()                             #超链接列号
        self.hyperlink_column.set(hyperlink_column_current_value)
        self.testinfo_download = tkinter.StringVar()                            #测试信息(下载表)
        self.testinfo_download.set(testinfo_download_current_value)
        self.testinfo_plan = tkinter.StringVar()                                #测试信息(计划表)
        self.testinfo_plan.set(testinfo_plan_current_value)
        self.object_plan = tkinter.StringVar()                                  #样品信息(计划表)
        self.object_plan.set(object_plan_current_value)
        self.addition_info=tkinter.StringVar()                                  #附加信息(计划表)
        self.addition_info.set(addition_info_current_value)
        self.test_data_folder_frame=None
        self.test_info_frame=None
        
        
      
        self.setup_ui()

    #显示或隐藏test_infor_frame
    def hide_frame1(self):                  
        if self.update_testinfo.get()==0:
            self.test_info_frame.grid_remove()
        else:
            self.test_info_frame.grid(row=3, column=1, sticky=W)
        
    #显示或隐藏test_data_folder_frame        
    def hide_frame2(self):
        if self.create_folder.get()==0:
            self.test_data_folder_frame.grid_remove()
            self.cb_with_foldername.grid_remove()
            self.if_download_spec.set(0)
            self.cb_download_spec.config(state=tkinter.DISABLED)
        else:
            self.test_data_folder_frame.grid(row=4, column=1, sticky=W)
            self.cb_with_foldername.grid(row=2,column=1,columnspan=3,stick=W)
            self.cb_with_foldername1.grid(row=2,column=2,columnspan=3,stick=W)
            self.cb_download_spec.config(state=tkinter.NORMAL)
    
        #增加文件夹
    def add_test_data_folder(self):
        i = len(self.test_data_folder_frame.winfo_children())
        s5 = tkinter.StringVar()
        entry5 = tkinter.Entry(self.test_data_folder_frame, width=25, textvariable=s5)
        entry5.grid(row=i + 2, column=1, columnspan=2)
        StringVarlist5.append(s5)

    #删除文件夹
    def delete_test_data_folder(self):
        childrenlist = self.test_data_folder_frame.winfo_children()
        i = len(self.test_data_folder_frame.winfo_children())
        if childrenlist[i - 1].winfo_class() == 'Entry':
            childrenlist[i - 1].destroy()
            StringVarlist5.pop()


            #下载inquired状态的test信息
    def download_inquired_test(self): 
        global driver
        global edge_options
        initial_browser(search)
        #driver.maximize_window()
        try:
            wait_element_clickable((By.XPATH, search_load),'load',30)
            time.sleep(7)  #
            showinfo(generate_report.report_status,'点击load',"black")
            el_load=driver.find_element(By.XPATH,search_load)
            try:
                el_load.click()
            except ElementNotInteractableException as e:
                driver.execute_script("arguments[0].scrollIntoView();", el_load)
                time.sleep(2)
                el_load.click() 
            time.sleep(1)
            showinfo(generate_report.report_status,'点击Test',"black")
            driver.find_element(By.XPATH,search_test).click() #使用text字段匹配方式定位 
            time.sleep(1)
            showinfo(generate_report.report_status,'点击预设的搜索条件',"black")
            driver.find_element(By.XPATH, '//*[text()="' + self.search_test_criteria.get() + '"]').click()  # 精确查找
            time.sleep(2)
            showinfo(generate_report.report_status,'点击search items',"black")
            el_search_items=driver.find_element(By.XPATH,search_search_items)
            try:
                el_search_items.click()
            except ElementNotInteractableException as e:   
                driver.execute_script("arguments[0].scrollIntoView();", el_search_items)
                time.sleep(2)
                el_search_items.click() 
            time.sleep(2)
            wait_element_invisibility((By.XPATH,search_process),'search process',30)
            time.sleep(2)
            showinfo(generate_report.report_status,'点击export',"black")
            driver.find_element(By.XPATH,search_export).click()
            time.sleep(1)
            showinfo(generate_report.report_status,'点击excel(all row)',"black")
            driver.find_element(By.XPATH,search_excel).click()
            time.sleep(1)
            wait_element_invisibility((By.XPATH,export_process),'export process',30)
            time.sleep(5)
        except Exception as e:
            showinfo(generate_report.report_status,'抓取inqured test过程中出现异常',"red")
            tkinter.messagebox.showinfo('提示', f'抓取inqured test过程中出现异常,请重新点击生成')
            sys.exit(1)
        download_file = get_download_file(15)
        return download_file

    #获取需要更新的testid列表
    def get_no_exist_idlist(self,download_file):
        global wb1
        global ws1
        global wb2
        global ws2
        no_exist_idlist = []  # 需要新增test的ID列表
        data_beginning = int(self.begin_row.get())
        current_max_row = 0
        try:
            wb2 = load_workbook(download_file)
        except Exception as e:
            showinfo(generate_report.report_status,'打开下载测试信息表时发生错误','red')
            tkinter.messagebox.showinfo('提示', '打开下载测试信息表时发生错误')
            sys.exit(1)    
        ws2 = wb2['Sheet 1']
        ws2_idlist = []
        for j in range(ws2.max_row - 1):
            ws2_idlist.append((ws2['A'][j + 1]).value)
            
        if self.update_testinfo.get()==1 and self.cobo_register_type.get()!='项目' :  #如果需要更新测试计划表并且登记登记类型不是项目，则在测试计划表中去判断单号是否需要创建
            if is_file_locked(plan_file):
                showinfo(generate_report.report_status,'测试计划表被打开,请关闭以后再操作','red')
                tkinter.messagebox.showinfo('提示', '测试计划表被打开,请关闭以后再操作')             
                sys.exit(1)
            else:
                wb1 = load_workbook(plan_file) #打开测试计划表
            ws1 = wb1[self.sheet_name.get()] 
            ws1_idlist = []
            current_max_row=get_max_row(ws1)

            for i in range(current_max_row - data_beginning+1):
                ws1_idlist.append(ws1['B'][i + data_beginning - 1].value)  
     
            for ws2_id in ws2_idlist:
                if ws2_id not in ws1_idlist:
                    showinfo(generate_report.report_status,ws2_id + '需要添加',"black")
                    no_exist_idlist.append(ws2_id)  # 将需要添加的test id 添加入列表
        else:    #如果不需要更新测试计划表或者需要更新测试计划表但登记类型是项目类型，则需要在数据文件夹中去判断单号是否需要创建，此过程时间比前一种方式更长
            for ws2_id in ws2_idlist:
                if search_folder3(test_dir,ws2_id) is None:
                    no_exist_idlist.append(ws2_id)
                
        return no_exist_idlist

    #更新计划表信息
    def add_inquired_test(self,no_exist_idlist):
        global plan_file
        global wb1
        global ws1
        global wb2
        global ws2
        plan_max_row = 0
        data_beginning = int(self.begin_row.get())
        down_index = self.testinfo_download.get()
        plan_index = self.testinfo_plan.get()
        down_index_list = down_index.split('_')
        plan_index_list = plan_index.split('_')
        hyperlink_cls=self.hyperlink_column.get()


        if ws1==None:
            try:
                wb1 = load_workbook(plan_file) #打开测试计划表
            except Exception as e:
                showinfo(generate_report.report_status,'打开测试计划表异常',"red")
                tkinter.messagebox.showinfo('提示', '打开测试计划表异常')
                sys.exit(1)
            ws1 = wb1[self.sheet_name.get()] 
            
        plan_max_row=get_max_row(ws1)                                         
        down_max_row=get_max_row(ws2) 
        
        ws1_idlist=ws1['B']
        ws2_idlist=ws2['A']
        if self.addition_info.get()!='':
            addition_inforlist=self.addition_info.get().split('@')
        if self.cobo_register_type.get() == '试验':    
            for testid in no_exist_idlist:
                for i,ws2_id in enumerate(ws2_idlist,start=1):
                    if testid == ws2_id.value:
                        ws1.cell(plan_max_row + 1, 1).value = plan_max_row-data_beginning+2
                        ws1.cell(plan_max_row + 1, 2).value = ws2.cell(i,1).value
                        if self.cobo_test_type.get() == '单个':
                            test_folder_path=search_folder2(test_dir_unc,testid)
                        else:
                            path=search_folder2(test_dir_unc,testid.split('.')[0])
                            test_folder_path=search_folder2(path,testid)
                        ws1.cell(plan_max_row + 1, int(hyperlink_cls)).hyperlink = test_folder_path  # ID 一栏建立超链接
                        for k in range(len(down_index_list)):
                            ws1.cell(plan_max_row + 1, int(plan_index_list[k])).value = ws2.cell(i,int(down_index_list[k])).value
                        if self.addition_info.get()!='':
                            for addition_infor in addition_inforlist:
                                ws1.cell(plan_max_row+1,int(addition_infor.split('_')[0])).value=addition_infor.split('_')[1]
                        showinfo(generate_report.report_status,testid + '测试信息添加成功',"black")
                        plan_max_row = plan_max_row + 1
        else:
            sequenceid_list=[]
            for testid in no_exist_idlist:
                sequenceid=testid[:10]
                sequenceid_list.append(sequenceid)
            unique_sequenceid_list=list(set(sequenceid_list))  #将sequenceid_list数组去除重复值
            
            for sequenceid in unique_sequenceid_list:
                sequenceid_existence=0
                for j,ws1_id in enumerate(ws1_idlist):
                    if sequenceid==ws1_id.value:
                        sequenceid_existence=1
                        break
                if sequenceid_existence==0:
                    for o,ws2_id in enumerate(ws2_idlist,start=1):
                        if ws2_id.value.split('.')[0]==sequenceid:
                            ws1.cell(plan_max_row+1,1).value=plan_max_row-data_beginning+2
                            ws1.cell(plan_max_row+1,2).value=sequenceid
                            test_folder_path=search_folder2(test_dir_unc,sequenceid)
                            ws1.cell(plan_max_row + 1, int(hyperlink_cls)).hyperlink = test_folder_path  # ID 一栏建立超链接
                            for k in range(len(down_index_list)):
                                if plan_index_list[k]!='2':          #前面已经赋值sequenceid
                                    ws1.cell(plan_max_row + 1, int(plan_index_list[k])).value = ws2.cell(o,int(down_index_list[k])).value
                            if self.addition_info.get()!='':
                                for addition_infor in addition_inforlist:
                                    ws1.cell(plan_max_row+1,int(addition_infor.split('_')[0])).value=addition_infor.split('_')[1]
                            showinfo(generate_report.report_status,sequenceid + '测试信息添加成功',"black")
                            plan_max_row = plan_max_row + 1
                            break
                    
    #创建测试文件夹                
    def create_test_folder(self,no_exist_idlist):  
        for testid in no_exist_idlist:
            test_folder=None
            for i in range(ws2.max_row):
                if ws2[f'A{i+1}'].value==testid:
                    project_name=ws2[f'B{i+1}'].value
                    test_name=ws2[f'C{i+1}'].value
                    #如果名称里含有非法字符则替代
                    if(bool(re.search(illegal_chars,project_name))):
                        project_name=re.sub(illegal_chars,'@',project_name)
                    if(bool(re.search(illegal_chars,test_name))):
                        test_name=re.sub(illegal_chars,'@',test_name)
            #test_dir = stringvar23.get()
            if self.cobo_test_type.get() == '序列':   #如果是序列方式则先创建项目文件夹
                if self.with_foldername.get()==1:
                    parent_test_folder = test_dir + '\\' + testid.split('.')[0]+'_'+project_name     #项目文件夹带项目名称
                else:    
                    parent_test_folder = test_dir + '\\' + testid.split('.')[0]                       #项目文件夹不带项目名称
                if self.with_foldername1.get()==1:
                    test_folder=parent_test_folder + '\\' + testid+'_' + test_name                   #测试文件夹带测试名称
                else:    
                    test_folder=parent_test_folder + '\\' + testid                                    #测试文件夹不带测试名称

                if not os.path.exists(parent_test_folder):
                    os.mkdir(parent_test_folder)
                else:
                    showinfo(generate_report.report_status,f'{parent_test_folder}文件夹已经存在',"black")    
            else:                            #如果非序列方式则直接创建测试文件夹
                if self.with_foldername1.get()==1:
                    test_folder=  test_dir + '\\' + testid+'_' + test_name
                else:    
                    test_folder= test_dir + '\\' + testid

            subfolder1 = test_folder + '\\' + '01_Test Data'
            subfolder2 = test_folder + '\\' + '02_Test Report'
            subfolder3 = test_folder + '\\' + 'Original data'
            #创建测试文件夹
            if not os.path.exists(test_folder):
                os.mkdir(test_folder)
            else:
                showinfo(generate_report.report_status,f'{test_folder}文件夹已经存在',"black")    
            #创建测试子文件夹1
            if not os.path.exists(subfolder1):
                os.mkdir(subfolder1)
            else:
                showinfo(generate_report.report_status,f'{subfolder1}文件夹已经存在',"black")
            #创建测试子文件夹2
            if not os.path.exists(subfolder2):
                os.mkdir(subfolder2)
            else:
                showinfo(generate_report.report_status,f'{subfolder2}文件夹已经存在',"black")
            #创建测试子文件夹3
            if not os.path.exists(subfolder3):
                os.mkdir(subfolder3)
            else:
                showinfo(generate_report.report_status,f'{subfolder3}文件夹已经存在',"black")    
            #创建自定义文件夹
            for s in StringVarlist5:
                folder = subfolder1 + '\\' + s.get()
                if not os.path.exists(folder):
                    os.mkdir(folder)
                else:
                    showinfo(generate_report.report_status,f'{folder}文件夹已经存在',"black")

            #创建link文件
            doc_address=Document()
            paragraph=doc_address.add_paragraph()
            run=paragraph.add_run()
            #unc_path=test_dir_driver_letter_unc_path+'\\'+test_folder.split(':',1)[1]
            unc_path=test_dir_unc+'\\'+test_folder.split(':',1)[1]
            add_hyperlink(paragraph,'ctr+鼠标左键点击这里访问测试数据',unc_path)
            doc_address.save(test_folder+'\\Raw data file link.docx')
            #创建test log文件
            test_log_text=open(test_folder+'\\test log.txt','w')
            test_log_text.close()
            showinfo(generate_report.report_status,f'{testid}测试文件夹创建成功',"black")

    #更新测试样品号码
    def update_sample_No(self,testidlist):
        global plan_file
        sample_no_lists=self.get_sample_id(testidlist)
        plan_max_row = 0
        data_beginning = int(self.begin_row.get())
        plan_object_index = self.object_plan.get()                                        
        plan_max_row=get_max_row(ws1)                                         
        for sample_no_list in sample_no_lists:
            for j in range(plan_max_row):
                if sample_no_list[0] == ws1.cell(j + data_beginning, 2).value:
                    ws1.cell(j + data_beginning, int(plan_object_index)).value = sample_no_list[1]  # 样品号码
                    showinfo(generate_report.report_status,f'{sample_no_list[0]}样品信息更新成功',"black")

    #获取测试样品号
    def get_sample_id(self,testidlist):  
        global driver
        global edge_options
        if driver.current_url!=search:
            initial_browser(search)
        sample_no_lists = []  #存放测试id,样品号码，样品个数的数组
        try:
            wait_element_clickable((By.XPATH, search_load),'load',30)
            time.sleep(3)
            showinfo(generate_report.report_status,'点击load',"black")
            driver.find_element(By.XPATH,search_load).click()
            time.sleep(1)
            showinfo(generate_report.report_status,'点击Test object',"black")
            driver.find_element(By.XPATH,search_test_object).click() #使用text字段匹配方式定位
            time.sleep(1)
            showinfo(generate_report.report_status,'点击预设的搜索条件',"black")
            driver.find_element(By.XPATH, '/html/body/div[11]/div//*[text()="' + self.search_object_criteria.get() + '"]').click()  # 精确查找
            time.sleep(2)                  
            el_add = driver.find_element(By.XPATH,search_add)
            showinfo(generate_report.report_status,'添加待搜索testid',"black")
            time.sleep(1)
            el_add.click()  # 点击add
            el_test = driver.find_element(By.XPATH, search_add_test)  # 获取test元素
            ActionChains(driver).move_to_element(el_test).perform()
            time.sleep(2)
            driver.find_element(By.XPATH, search_add_test_id).click()  # 获取test下的ID并点击
            el_input = driver.find_element(By.XPATH,f'/html/body/div[5]/div/div[2]/div/div/div[2]/div/div/div/div[4]/div/div[2]/div[2]/table/tbody/tr[{1}]/td[4]/div/div/input')
            for testid in testidlist:    
                try:
                    # 获取输入框标签                     
                    ActionChains(driver).double_click(el_input).perform()  # 双击输入框，消除输入框自带默认值
                    el_input.clear()
                    time.sleep(2)
                    el_input.send_keys(testid)
                    showinfo(generate_report.report_status,'点击search items',"black")
                    driver.find_element(By.XPATH,search_search_items).click()  # 点击search
                    time.sleep(2)
                    wait_element_invisibility((By.XPATH,search_process),'search process',30)
                    time.sleep(2)
                    showinfo(generate_report.report_status,'点击export',"black")
                    driver.find_element(By.XPATH,search_export).click()
                    time.sleep(1)
                    showinfo(generate_report.report_status,'点击excel(all row)',"black")
                    driver.find_element(By.XPATH,search_excel).click()  # 点击excel(all row)
                    time.sleep(1)
                    wait_element_invisibility((By.XPATH,export_process),'export process',30)
                    #点击filter回到search界面
                    driver.find_element(By.XPATH,search_filter).click()
                    download_file = get_download_file(15)
                    if download_file!=None:
                        wb = load_workbook(download_file)
                        ws = wb['Sheet 1']
                        showinfo(generate_report.report_status,f'{testid}样品信息表下载成功',"black")
                        sample_no=''                                    #样品号码
                        sample_quantity=ws.max_row-1                    #样品数量
                        for j in range(ws.max_row-1):
                            if sample_no=='':
                                sample_no=str(ws.cell(j + 2, 3).value)
                            else:    
                                sample_no = sample_no + ',' + str(ws.cell(j + 2, 3).value)
                        sample_no_list=[]
                        sample_no_list.append(testid)             
                        sample_no_list.append(sample_no)
                        sample_no_list.append(sample_quantity)    
                        sample_no_lists.append(sample_no_list)
                        wb.close()
                        driver.find_element(By.XPATH,search_filter).click()
                except Exception as e:
                    showinfo(generate_report.report_status,f'抓取{testid}样品号码过程中出现异常,异常描述:{str(e)}',"red")
                    driver.find_element(By.XPATH,search_filter).click()          #点击filter回到search界面
        except Exception as e:
            showinfo(generate_report.report_status,f'抓取样品号码过程中出现异常,异常描述:{str(e)}',"red")
            driver.find_element(By.XPATH,search_filter).click()
            #点击filter回到search界面
        return sample_no_lists    


    #下载测试规范        
    def download_spec(self,testidlist):
        global driver
        for testid in testidlist:
            try:
                el_testid = driver.find_element(By.XPATH, '//*[text()="' + testid + '"]')  # 精确查找
                time.sleep(1)
                ActionChains(driver).double_click(el_testid).perform()
                wait_element_clickable((By.XPATH,search_download_spec),'download',30)
                time.sleep(2)
                driver.find_element(By.XPATH,search_download_spec).click()  # 点击download
                time.sleep(3)
                download_file = get_download_file(15)
                if download_file is not None:
                    showinfo(generate_report.report_status,f'{testid}测试规范下载成功',"black")
                    test_folder_path = search_folder3(test_dir, testid)
                    if test_folder_path is not None:
                        move_file(download_file, test_folder_path,1)
                        showinfo(generate_report.report_status,f'{testid}测试规范移动成功',"black")
                    else:
                        showinfo(generate_report.report_status,f'{testid}测试文件夹不存在',"red")
                else:
                    showinfo(generate_report.report_status,f'超时,{testid}测试规范下载失败',"red")
                el_close_window = driver.find_element(By.XPATH,search_spec_close_window)
                ActionChains(driver).move_to_element(el_close_window).perform()
                time.sleep(1)
                el_close_window.click()
            except Exception as e:
                if isElementExist(search_download_spec_cancel):  #如果下载页面未关闭，则关闭
                    driver.find_element(By.XPATH,search_download_spec_cancel).click()
                    time.sleep(1)
                showinfo(generate_report.report_status,f'下载{testid}规范过程中出现异常,异常描述:{str(e)}',"red")
        #点击filter回到search界面
        driver.find_element(By.XPATH,search_filter).click()       
        #更新测试计划表
    def update_test_matrix(self): 
        global plan_file
        global plan_file_unc
        global test_dir_unc
        global wb1
        global ws1
        global wb2
        global ws2
        plan_file_drive_letter=plan_file[0]  #获取计划表地址的盘符
        test_dir_driver_letter=test_dir[0]   #获取测试文件夹根目录盘符
        plan_file_drive_letter_unc_path=get_unc_path(plan_file_drive_letter)
        test_dir_driver_letter_unc_path=get_unc_path(test_dir_driver_letter)
        if plan_file_drive_letter_unc_path is not None:
            plan_file_unc=plan_file_drive_letter_unc_path+'//'+plan_file.split(':',1)[1]         #计划表unc地址
            print('1')
            print(plan_file_unc)
        else:
            plan_file_unc=plan_file
            print('2')
            print(plan_file_unc)
        if test_dir_driver_letter_unc_path is not None:
            test_dir_unc=test_dir_driver_letter_unc_path+'//'+test_dir.split(':',1)[1]           #数据文件夹unc地址
            print('3')
            print(test_dir_unc)
        else:    
            test_dir_unc=test_dir
            print('4')
            print(test_dir_unc)
        if self.update_testinfo.get()==1:
            if is_file_locked(plan_file): 
                showinfo(generate_report.report_status,'测试计划表被打开,请关闭以后再操作','red') 
                tkinter.messagebox.showinfo('提示', '测试计划表被打开,请关闭以后再操作')
                sys.exit(1)
        if self.update_testinfo.get()==1 or self.create_folder.get()==1 or self.if_download_spec.get()==1:
            showinfo(generate_report.report_status,'正在下载inquired测试信息表',"black")
            try:
                download_file = self.download_inquired_test()
            except Exception as e:
                showinfo(generate_report.report_status,f'下载测试信息表出错{str(e)}',"red")
                tkinter.messagebox.showinfo('提示', '下载测试信息表出错')    
                sys.exit(1)
            if download_file is not None:
                showinfo(generate_report.report_status,'inquired测试信息表下载成功',"black")
                showinfo(generate_report.report_status,'正在获取需要创建的测试单号...',"black")
                no_exist_idlist = self.get_no_exist_idlist(download_file)  #获取需要更新的testid
                showinfo(generate_report.report_status,f'待新增的testid列表:{no_exist_idlist}',"black") 
                if len(no_exist_idlist) == 0:
                    showinfo(generate_report.report_status,'计划表无需更新',"green")
                    tkinter.messagebox.showinfo('提示', '计划表无需更新')
                else:
                    if self.create_folder.get()==1:    #如果勾选了创建测试文件夹
                        showinfo(generate_report.report_status,'正在创建测试文件夹',"black")
                        try:
                            self.create_test_folder(no_exist_idlist)
                            showinfo(generate_report.report_status,'所有测试文件夹创建完毕',"green")
                        except Exception as e:
                            showinfo(generate_report.report_status,f'创建测试文件夹过程中发生异常{e}',"red")
                            tkinter.messagebox.showinfo('提示', '创建测试文件夹过程中发生异常')
                            sys.exit(1)
                    if self.if_download_spec.get()==1:     #如果勾选了下载测试规范
                        showinfo(generate_report.report_status,'正在下载测试规范',"black")
                        self.download_spec(no_exist_idlist)  # 下载需要新增test的测试规范
                        showinfo(generate_report.report_status,'所有测试规范下载完毕',"green")      
                    else: #如果不需要下载规范必须回到搜索条件页面
                        driver.find_element(By.XPATH,search_filter).click()
                    if self.update_testinfo.get()==1:     #如果勾选了更新测试信息表
                        showinfo(generate_report.report_status,'正在更新测试基本信息',"black")
                        self.add_inquired_test(no_exist_idlist)
                        showinfo(generate_report.report_status,'测试信息更新完毕',"green")
                        if self.cobo_register_type.get() == '试验':  
                            showinfo(generate_report.report_status,'正在更新样品信息',"black")
                            #点击filter回到search界面
                            driver.find_element(By.XPATH,search_filter).click()
                            self.update_sample_No(no_exist_idlist)  # 添加样品号码
                            showinfo(generate_report.report_status,'样品信息更新完毕',"green")
                    try:
                        if self.update_testinfo.get()==1:  #如果勾选了更新测试信息表,则需要保存测试计划表
                            wb1.save(plan_file)
                            #shared_pyxl_save(plan_file,wb1)
                            ws1=None        #修改ws1值必须 声明global ws1
                            wb1.close()
        
                        tkinter.messagebox.showinfo('提示', '所有更新完毕')
                    except PermissionError as e:
                        showinfo(generate_report.report_status,'测试计划表文件被打开无法被保存,请手动关闭文件',"red")
                        tkinter.messagebox.showinfo('提示', '测试计划表文件被打开无法被保存,请手动关闭文件再重试一次')
            else:
                showinfo(generate_report.report_status,'超时,测试信息表下载失败',"red")
        else:
            showinfo(generate_report.report_status,'请至少选择一个更新类型',"red")
            tkinter.messagebox.showinfo('提示', '请至少选择一个更新类型')

    def setup_ui(self):
        # plan_frame
        plan_frame = tkinter.LabelFrame(self.root, text='计划表更新',font='Arial 12 bold italic')
        plan_frame.grid(row=2, column=1, columnspan=3, sticky=W)
        #basic_frame
        basic_frame = tkinter.LabelFrame(plan_frame, text='一般参数设置',font='Arial 12 bold italic')
        basic_frame.grid(row=1, column=1,sticky=W)
        tkinter.Label(basic_frame, text="试验类型", width=8).grid(row=1, column=2)
        self.cobo_test_type = ttk.Combobox(basic_frame, width=5)
        self.cobo_test_type['value'] = ('单个', '序列')
        self.cobo_test_type.current(cobo_test_type_current_value)
        self.cobo_test_type.grid(row=1, column=3)

        tkinter.Label(basic_frame, text="登记类型", width=8).grid(row=1, column=4)
        self.cobo_register_type = ttk.Combobox(basic_frame, width=5)
        self.cobo_register_type['value'] = ('试验', '项目')
        self.cobo_register_type.current(cobo_register_type_current_value)
        self.cobo_register_type.grid(row=1, column=5)
        #update_item_frame
        update_item_frame=tkinter.LabelFrame(plan_frame,text='更新类型选择',font='Arial 12 bold italic')
        update_item_frame.grid(row=2,column=1,sticky=W)
        tkinter.Checkbutton(update_item_frame, text='更新测试信息表', variable=self.update_testinfo, onvalue=1, offvalue=0,command=self.hide_frame1).grid(row=1, column=1, sticky=W)
        tkinter.Checkbutton(update_item_frame, text='创建测试文件夹', variable=self.create_folder, onvalue=1, offvalue=0,command=self.hide_frame2).grid(row=1, column=2, sticky=W)

        self.cb_download_spec = tkinter.Checkbutton(update_item_frame, text='下载测试规范', variable=self.if_download_spec, onvalue=1, offvalue=0)
        self.cb_download_spec.grid(row=1, column=3, sticky=W)
        self.cb_with_foldername = tkinter.Checkbutton(update_item_frame, text='项目文件夹名称', variable=self.with_foldername, onvalue=1, offvalue=0)
        self.cb_with_foldername.grid(row=2,column=1, sticky=W)
        self.cb_with_foldername1 = tkinter.Checkbutton(update_item_frame, text='测试文件夹名称', variable=self.with_foldername1, onvalue=1, offvalue=0)
        self.cb_with_foldername1.grid(row=2,column=2, sticky=W)
        #test_info_frame
        self.test_info_frame = tkinter.LabelFrame(plan_frame, text='测试信息更新',font='Arial 12 bold italic')
        self.test_info_frame.grid(row=3, column=1, sticky=W)
        #search_keyword_frame
        search_keyword_frame=tkinter.LabelFrame(self.test_info_frame,text='搜索字段设置',font='Arial 12 bold italic')
        search_keyword_frame.grid(row=1, column=1, columnspan=5,sticky=W)
        tkinter.Label(search_keyword_frame, text="测试信息", width=10).grid(row=1, column=1,stick=W)
        tkinter.Entry(search_keyword_frame, width=12, textvariable=self.search_test_criteria).grid(row=1, column=2,stick=W)
        tkinter.Label(search_keyword_frame, text="样品信息", width=10).grid(row=1, column=3,stick=W)
        tkinter.Entry(search_keyword_frame, width=12, textvariable=self.search_object_criteria).grid(row=1, column=4,stick=W)
        tkinter.Label(search_keyword_frame, text="测试信息2", width=10).grid(row=1, column=5,stick=W)
        tkinter.Entry(search_keyword_frame, width=12, textvariable=self.search_test_criteria1).grid(row=1, column=6,stick=W)


        #relationship_frame
        relationship_frame=tkinter.LabelFrame(self.test_info_frame,text='表格映射',font='Arial 12 bold italic')
        relationship_frame.grid(row=2,column=1,stick=W)
        tkinter.Label(relationship_frame, text="计划表分页", width=10).grid(row=1, column=1)
        tkinter.Entry(relationship_frame, width=12, textvariable=self.sheet_name).grid(row=1, column=2)
   
        tkinter.Label(relationship_frame, text="计划表起始行", width=10).grid(row=1, column=3)
        tkinter.Entry(relationship_frame, width=12, textvariable=self.begin_row).grid(row=1, column=4,sticky=W)
        tkinter.Label(relationship_frame, text="超链接列号", width=10).grid(row=1, column=5,sticky=W)
        tkinter.Entry(relationship_frame, width=12, textvariable=self.hyperlink_column).grid(row=1, column=6,sticky=W)

        tkinter.Label(relationship_frame, text='测试信息').grid(row=2, column=2)
        tkinter.Label(relationship_frame, text='样品号码').grid(row=2, column=3)
        tkinter.Label(relationship_frame, text='下载信息表').grid(row=4, column=1)
        tkinter.Label(relationship_frame, text='测试计划表').grid(row=3, column=1)
        tkinter.Label(relationship_frame, text='附加信息').grid(row=2, column=4)

        tkinter.Entry(relationship_frame, width=12, textvariable=self.testinfo_download).grid(row=4, column=2)  # 测试信息(下载表)
        tkinter.Entry(relationship_frame, width=12, textvariable=self.testinfo_plan).grid(row=3, column=2)  # 测试信息(计划表)
        tkinter.Entry(relationship_frame, width=12, textvariable=self.object_plan).grid(row=3, column=3)  # 样品信息(计划表)
        tkinter.Entry(relationship_frame, width=25, textvariable=self.addition_info).grid(row=3, column=4,columnspan=2)  # 附加信息

        #test_data_folder_frame
        self.test_data_folder_frame = tkinter.LabelFrame(plan_frame, text='自定义数据文件夹',font='Arial 12 bold italic')
        self.test_data_folder_frame.grid(row=4, column=1, sticky=W)
        ttk.Button(self.test_data_folder_frame, text="添加", command=self.add_test_data_folder).grid(row=1, column=1, sticky=W)
        ttk.Button(self.test_data_folder_frame, text="删除", command=self.delete_test_data_folder).grid(row=1, column=2, sticky=E)
        for i in range(len(folder_list)):
            s5 = tkinter.StringVar()
            s5.set(folder_list[i])
            entry5 = tkinter.Entry(self.test_data_folder_frame, width=25, textvariable=s5)
            entry5.grid(row=i + 2, column=1, columnspan=2) 
            StringVarlist5.append(s5)

        ttk.Button(plan_frame, text="更新", command=lambda:run_thread(self.update_test_matrix,'update_test_matrix')).grid(row=5, column=1, sticky=W)
        self.hide_frame1()  #
        self.hide_frame2()  #

class TestCurve:
    def __init__(self, root):
        self.root = root
        # 变量
        self.data_address = tkinter.StringVar()
        self.data_character = tkinter.StringVar()               #数据特征
        self.data_character.set(data_character_current_value)
        self.sampling_frequency = tkinter.StringVar()               #采样频率
        self.sampling_frequency.set(sampling_frequency_current_value) 
        self.x_unit = tkinter.StringVar()               #X轴单位
        self.x_unit.set(x_unit_current_value)
        self.coordinate_typle = tkinter.IntVar()                     #坐标系类型
        self.coordinate_typle.set(coordinate_typle_current_value) 
        self.curve_config_frame=None
        self.setup_ui()
        
    
    def setup_ui(self):
        address_frame1 = tkinter.LabelFrame(self.root, text='地址选择', width=550,font='Arial 12 bold italic')
        address_frame1.grid(row=1, column=1, columnspan=5, sticky=W)
        ttk.Button(master=address_frame1,width=15,text='选择数据地址',command=self.select_data_dir).grid(row=1, column=1,sticky=W)
        tkinter.Message(address_frame1, textvariable=self.data_address, width=1000).grid(row=1, column=2, columnspan=10, rowspan=2,sticky=W)
        ttk.Button(address_frame1,width=15,text='生成曲线',command=lambda:run_thread(self.generate_curve(self.data_address.get(),0),'generate_curve')).grid(row=3, column=1,sticky=W)
        self.curve_config_frame=tkinter.LabelFrame(self.root, text='曲线参数设置', width=500,font='Arial 12 bold italic')
        self.curve_config_frame.grid(row=2, column=1, columnspan=5, sticky=W)
        tkinter.Label(self.curve_config_frame, text='数据特征').grid(row=1, column=1,sticky=W)
        tkinter.Entry(self.curve_config_frame, width=25, textvariable=self.data_character).grid(row=1, column=2,columnspan=2,sticky=W)

        tkinter.Label(self.curve_config_frame, text='采样频率').grid(row=2, column=1,sticky=W)
        tkinter.Entry(self.curve_config_frame, width=25, textvariable=self.sampling_frequency).grid(row=2, column=2,columnspan=2,sticky=W)

        tkinter.Label(self.curve_config_frame, text='x轴单位').grid(row=3, column=1,sticky=W)
        tkinter.Entry(self.curve_config_frame, width=25, textvariable=self.x_unit).grid(row=3, column=2,columnspan=2,sticky=W)
        
        
        radio71=tkinter.Radiobutton(self.curve_config_frame,text='单独坐标系',variable=self.coordinate_typle,value=1)
        radio72=tkinter.Radiobutton(self.curve_config_frame,text='混合坐标系',variable=self.coordinate_typle,value=2)
        radio73=tkinter.Radiobutton(self.curve_config_frame,text='复合坐标系',variable=self.coordinate_typle,value=3)
        if self.coordinate_typle.get()==1:
            radio71.select()
        if self.coordinate_typle.get()==2:
            radio72.select()    
        if self.coordinate_typle.get()==3:
            radio73.select()     
        radio71.grid(row=4, column=1)
        radio72.grid(row=4, column=2)
        radio73.grid(row=4, column=3)


        tkinter.Label(self.curve_config_frame, text='数据列特征').grid(row=5, column=1,sticky=W)
        ttk.Button(self.curve_config_frame, text="添加", command=self.add_column_feature).grid(row=5, column=2, sticky=W)
        ttk.Button(self.curve_config_frame, text="删除", command=self.delete_column_feature).grid(row=5, column=3, sticky=W)
        for i in range(len(column_list)):
            s6 = tkinter.StringVar()
            s6.set(column_list[i])
            entry6 = tkinter.Entry(self.curve_config_frame, width=40, textvariable=s6)
            entry6.grid(row=i + 6, column=1, columnspan=3,sticky=W) 
            StringVarlist6.append(s6)

        #选择数据目录
    def select_data_dir(self):  
        path = tkinter.filedialog.askdirectory().replace('/','\\')
        if path!='':
            self.data_address.set(path)    

        #增加数据列特征
    def add_column_feature(self):
        i = len(self.curve_config_frame.winfo_children())
        s6 = tkinter.StringVar()
        entry6 = tkinter.Entry(self.curve_config_frame, width=40, textvariable=s6)
        entry6.grid(row=i + 6, column=1, columnspan=3,sticky=W)
        StringVarlist6.append(s6)

    #删除数据列特征
    def delete_column_feature(self):
        childrenlist = self.curve_config_frame.winfo_children()
        i = len(self.curve_config_frame.winfo_children())
        if childrenlist[i - 1].winfo_class() == 'Entry':
            childrenlist[i - 1].destroy()
            StringVarlist6.pop()

    def generate_curve(self,test_data_path,type):
        # test_data_path=self.data_address.get()
        string_value=self.data_character.get()
        test_datalist=search_file3(test_data_path,string_value)
        column_list=[]
        for s in StringVarlist6:
            column_list.append(s.get())
        for test_data in test_datalist:
            data_name=os.path.basename(test_data)
            pic=self.draw_figure(test_data,column_list)
            showinfo(generate_report.report_status,f'{data_name}曲线生成完毕','green')
            if type==1:  #如果是1,则代表这是自动生成报告过程中的曲线生成
                root_dir=search_root_dir(test_data)
                move_file(pic,root_dir,0)

    #生成曲线
    def draw_figure(self,file_path,column_list):
        import matplotlib.pyplot as plt
        from matplotlib.ticker import AutoMinorLocator, AutoLocator
        import pandas as pd
        import numpy as np
            # 读取数据文件（假设制表符分隔）
        df = pd.read_csv(file_path, sep='\t', header=0)
        df.fillna(method='ffill',inplace=True)
        # 提取时间列（假设第一列是时间）
        time = df.iloc[:, 0]
        x_values=np.arange(1,len(df)+1)
        cols1=[]                           #将列名不分组存放入cols1
        cols2=[]                           #将列名分组存放入cols2
        for column in column_list:
            cols3=[]
            cols3.append(column)           #将列名放入cols3第一位
            for col in df.columns:
                if column in col:
                    cols1.append(col)
                    cols3.append(col)
            cols2.append(cols3)     


        valid_columns = df.columns.tolist()
        num_columns=len(cols1)
        #num_columns = len(current_cols)+len(voltage_cols)+len(temperature_cols)
        # print(num_columns)
        column_names = valid_columns[1:]  # 获取除 'Time' 列以外的列名
        # print(column_names)

        # 图形设置
        width = 12
        height = width * (3 / 4)
        #fig, axs = plt.subplots(nrows=num_columns, ncols=1, figsize=(width, height), sharex=True)
        #颜色列表
        colors = [
            '#FF5733', '#33FF57', '#3357FF', '#FF33A1', '#FFDB33',
            '#33FFD5', '#A133FF', '#FF8C33', '#33FF8C', '#FFC300'
        ]
        sampling_rate=self.sampling_frequency.get()
        t=1/int(sampling_rate)
        unit=self.x_unit.get()
        if self.coordinate_typle.get()==1:
            fig, axs = plt.subplots(nrows=num_columns, ncols=1, figsize=(width, height), sharex=True)
            for plot_index, col in enumerate(cols1):
                axs[plot_index].plot(x_values, df[col], label=col,color = colors[plot_index])
                axs[plot_index].set_ylabel(col)
                axs[plot_index].yaxis.set_major_locator(AutoLocator()) 
                axs[plot_index].yaxis.set_minor_locator(AutoMinorLocator())
                axs[plot_index].legend(loc="upper left")
                axs[plot_index].grid(True, alpha=0.3)
            axs[num_columns - 1].set_xlabel(f'Time({str(t)+unit})',loc='left') 
        if self.coordinate_typle.get()==2:
            plot_index=0
            fig, ax = plt.subplots(figsize=(12, 8))
            fig.subplots_adjust(right=0.75)  # 为右侧坐标轴预留空间
            lines=[]
            labels=[]
            for i,col in enumerate(cols2):
                if i==0:
                    ax1=ax
                else:
                    ax1 = ax.twinx()
                    ax1.spines['right'].set_position(('outward', 60*i))
                ax1.set_ylabel(col[0])
                for j,c in enumerate(col[1:]):
                    ax1.plot(x_values, df[c].values, color=colors[plot_index], linestyle='-', linewidth=1.5, alpha=0.5,label=c)
                    ax1.yaxis.set_major_locator(AutoLocator()) 
                    ax1.yaxis.set_minor_locator(AutoMinorLocator())
                    plot_index=plot_index+1
                line,label=ax1.get_legend_handles_labels()   
                lines=lines+line
                labels=labels+label    
                ax1.grid(True, alpha=0.3)    
            ax.set_xlabel(f'Time({str(t)+unit})',loc='left')
            ax.legend(lines,labels,loc='upper center', bbox_to_anchor=(0.5, -0.15),ncol=3, fontsize=10)
        if self.coordinate_typle.get()==3:
            plot_index=0
            fig, axs = plt.subplots(nrows=len(cols2), ncols=1, figsize=(width, height), sharex=True)
            for i,col in enumerate(cols2):
                lines=[]
                labels=[]
                axs[i].set_ylabel(col[0])    
                for j,c in enumerate(col[1:]):
                    axs[i].plot(x_values, df[c].values, color=colors[plot_index], linestyle='-', linewidth=1.5, alpha=0.5,label=c)
                    axs[i].yaxis.set_major_locator(AutoLocator()) 
                    axs[i].yaxis.set_minor_locator(AutoMinorLocator())
                    plot_index=plot_index+1 
                line,label=axs[i].get_legend_handles_labels()  
                lines=lines+line
                labels=labels+label 
                axs[i].legend(lines,labels,loc='upper center', bbox_to_anchor=(0.5, -0.15),ncol=3, fontsize=10)    
                axs[i].grid(True,alpha=0.3) 
            axs[len(cols2)-1].set_xlabel(f'Time({str(t)+unit})',loc='left')        
        plt.tight_layout()
        # plt.show()
        # 显示图形
        file_name,file_extension=os.path.splitext(file_path)
        pic=file_name+'.png'
        plt.savefig(pic, format='png', dpi=300, bbox_inches='tight')
        plt.close(fig)
        return pic        

class DateRangePicker:
    def __init__(self, root):
        self.root = root
        self.start_date = None
        self.end_date = None
        self.result = False  # 用户是否点击了确认
        self.object_deliverdate=None
    
    def show(self,date):
        """显示日期选择器并等待用户操作"""
        self.object_deliverdate=date
        self.select_date_window = tk.Toplevel(self.root)
        self.select_date_window.title("日期范围选择器")
        x,y=get_middle_position(self.root,self.select_date_window)  #获取居中位置坐标
        self.select_date_window.geometry(f"+{x}+{y}")
        self.select_date_window.transient(self.root)
        self.select_date_window.grab_set()  # 模态
        # 防止用户关闭窗口
        self.select_date_window.protocol("WM_DELETE_WINDOW", self.on_close)
        # 设置界面
        self.setup_ui()
        # 居中显示
        #center_window(self.select_date_window)
        # 等待窗口关闭
        self.root.wait_window(self.select_date_window)
        # 返回选择结果
        return self.result, self.start_date, self.end_date

    def setup_ui(self):
        """设置界面"""
        # 标题
        ttk.Label(self.select_date_window, text="日期范围选择", font=("Arial", 16, "bold")).grid(row=1,column=1,columnspan=2,pady=10)
        # 选择开始日期按钮
        self.start_btn = ttk.Button(self.select_date_window, text="选择开始日期",command=lambda: self.select_date("start"))
        self.start_btn.grid(row=2,column=1,columnspan=2,pady=10)
        # 开始日期显示
        self.start_label = ttk.Label(self.select_date_window, text="未选择", font=("Arial", 10), foreground="blue")
        self.start_label.grid(row=3,column=1,columnspan=2)
        # 选择结束日期按钮
        self.end_btn = ttk.Button(self.select_date_window, text="选择结束日期",command=lambda: self.select_date("end"))
        self.end_btn.grid(row=4,column=1,columnspan=2,pady=10)
        # 结束日期显示
        self.end_label = ttk.Label(self.select_date_window, text="未选择", font=("Arial", 10), foreground="green")
        self.end_label.grid(row=5,column=1,columnspan=2)
        # 清除按钮
        ttk.Button(self.select_date_window, text="清除", command=self.clear_selection).grid(row=6,column=1,padx=2)
        # 确认按钮
        ttk.Button(self.select_date_window, text="确认", command=self.submit_selection).grid(row=6,column=2,padx=2)          
        
    def select_date(self, date_type):
        def on_select():
            selected = cal.get_date()
            if date_type == "start":
                self.start_date = selected
                self.start_label.config(text=f"开始日期: {selected}")
            else:
                self.end_date = selected
                self.end_label.config(text=f"结束日期: {selected}")
            popup.destroy()
        """选择日期（开始或结束）"""
        popup = tk.Toplevel(self.select_date_window)
        popup.title(f"选择{('开始' if date_type == 'start' else '结束')}日期")
        x,y=get_middle_position(self.select_date_window,popup)
        popup.geometry(f"+{x}+{y}")
        popup.transient(self.select_date_window)
        popup.grab_set()
        # 创建日历
        today = datetime.now()
        cal = Calendar(popup, selectmode='day',year=today.year,month=today.month,day=today.day,date_pattern='yyyy-mm-dd')
        cal.pack(pady=20)
        ttk.Button(popup, text="确定",command=on_select).pack(pady=10)
        # 双击选择
        cal.bind("<Double-Button-1>",lambda e:on_select())
        
    def cancel_selection(self):
        """取消选择"""
        self.result = False
        self.select_date_window.destroy()       

    def on_close(self):
        """窗口关闭事件"""
        self.cancel_selection()

    def clear_selection(self):
        """清除选择"""
        self.start_date = None
        self.end_date = None
        self.start_label.config(text="未选择")
        self.end_label.config(text="未选择")

    def submit_selection(self):
        """确认选择"""
        if not self.start_date or not self.end_date:
            tkinter.messagebox.showwarning("警告", "请先选择开始和结束日期")
            return
        # 验证日期
        try:
            start = datetime.strptime(self.start_date, "%Y-%m-%d")
            end = datetime.strptime(self.end_date, "%Y-%m-%d")
            deliverydate=''
            if self.object_deliverdate!='':
                deliverydate=datetime.strptime(self.object_deliverdate, "%Y-%m-%d")
            if end < start:
                tkinter.messagebox.showwarning("警告", "结束日期不能早于开始日期")
                return
            if deliverydate!='' and deliverydate>start:
                tkinter.messagebox.showwarning("警告", "试验开始日期不能早于样品达到日期")
                return
            self.result = True
            self.select_date_window.destroy()
        except Exception as e:
            tkinter.messagebox.showerror("错误", f"日期格式错误: {str(e)}")
            print(str(e))

class EdateApi:
    def __init__(self):
        self.EDATE_USER = 'ZHS6CNG'
        self.EDATE_PASSWORD = 'tyjktz198558'
        self.API_GATEWAY_KEY ='b40a80fa-9eda-4f74-a489-16a7c5e0f6f5'
        self.BaseUrl_eDaTe_Prod = 'https://edate.webapp.inside.bosch.cloud/edate/api'
        self.BaseUrl_eDaTe_Prod_Gateway_BCN = 'https://ews-bcn.api.bosch.com/edate/api/v2'
        self.BaseUrl_eDaTe_Prod_Gateway_SL1 = 'https://ews-emea.api.bosch.com/edate/api/v2/'

        self.BaseUrl_eDaTe_Test = 'https://fe0vm7807.de.bosch.com:8443/edate/api'
        self.BaseUrl_eDaTe_Test_Gateway_BCN = 'https://ews-bcn.api.bosch.com/edate/api/test'
        self.BaseUrl_eDaTe_Test_Gateway_SL1 = 'https://ews-emea.api.bosch.com/edate/api/test'

        self.BaseUrl_eDaTe_Prod_GatewayFinal_SL2 = 'https://ews-esz-emea.api.bosch.com/smartlab/edate/api'
        self.BaseUrl_eDaTe_Prod_GatewayFinal_SL1 = 'https://ews-emea.api.bosch.com/smartlab/edate/api'
        self.API_Gateway_Key_SL1 = 'b40a80fa-9eda-4f74-a489-16a7c5e0f6f5'

        self.SML1CNG_Credentials = 'ZHS6CNG:tyjktz198558'
        #Auth_Credentials = {{SML1CNG_Credentials}}


      #ZHS6cng
    def get_test_orders_by_filter(self,filter_name,user_name):
        """获取测试订单的 EntityId 列表"""
        payload = {"name": filter_name,"folderName": f"generic-search-{user_name}","maxResults": 0}
        return self.call_edate_api(endpoint="/genericsearch/v1/TestInstance_UserSearch/execute-stored",payload=payload)

    def get_test_order_details(self,entity_id):
        """根据 EntityId 获取测试订单详情"""
        payload = {
            "filterCriteria": f"TestInstance#id={entity_id}",
            "resultFields": [
                {"fieldId": "TestInstance#id"},
                {"fieldId": "TestInstance#testInstanceId"},
                {"fieldId": "TestInstance#status"},
                {"fieldId": "TestInstance#plannedStart"},
                {"fieldId": "TestInstance#plannedEnd"},
                {"fieldId": "TestObject#status"},
                {"fieldId": "TestObject#deliveryDate"},
                {"fieldId": "Orderer#displayName"},
                {"fieldId": "Responsible#displayName"},

                # 其他需要的字段...
            ]
        }
        return self.call_edate_api(endpoint="/genericsearch/v1/TestInstance_UserSearch/search",payload=payload)

    def get_test_objects_by_entity_ids(self,entity_ids):
        """批量获取测试单关联的样品信息"""
        if isinstance(entity_ids, list):
            entity_ids = ",".join(map(str, entity_ids))
        
        payload = {
            "filterCriteria": f"TestInstance#id IN ({entity_ids})",
            "maxResults": 0,
            "resultFields": [
                {"fieldId": "TestObject#testObjectId"},
                {"fieldId": "TestInstance#id"},
                {"fieldId": "TestObject#deliveryDate"},
                # 其他需要的字段...
            ]
        }
        return self.call_edate_api(endpoint="/genericsearch/v1/TestObjectTestInstanceRelation/search",payload=payload)

    def call_edate_api(self,
        method="POST",
        url=None,
        endpoint="",
        base_url='https://edate.webapp.inside.bosch.cloud/edate/api',
        payload=None,
        gateway_key=None,
        auth=None
    ):
        """通用 eDaTe API 调用函数"""
        headers = {
            "Content-Type": "application/json;charset=UTF-8",
            "Accept": "application/json;charset=UTF-8",
        }
        if gateway_key:
            headers["KeyId"] = gateway_key
        if auth is None:
            auth = HTTPBasicAuth(self.EDATE_USER, self.EDATE_PASSWORD)
        full_url = f"{base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        response = requests.request(
            method,
            full_url,
            headers=headers,
            auth=auth,
            data=json.dumps(payload) if payload else None
        )
        if response.status_code != 200:
            raise Exception(f"API Error: {response.status_code} - {response.text}")
        return response.json()


#ERROR_BAD_EXE_FORMAT = 193      
register_name=os.getlogin()  #获取当前登录账号的用户名
edge_options = None  
driver = None
# driver = webdriver.Edge(executable_path="U:\python\msedgedriver.exe")

test_id = ''
report_dir = ''

excel_file_list=[]
data_dir=''
visit_times=0
#doc = Document()
doc = None
wb1 = None   #excel计划表
ws1 = None   #计划表sheet
wb2 = None   #excel下载表
ws2 = None   #下载表sheet

illegal_chars = r'[<>:"/\|?*]'    #非法字符     
needreport=1          #  1:需要报告，0：不需要报告

active_threadlist=['generate_test_report','update','upload_test_report','upload_data','upload_link','generate_curve','copy','cut','update_test_spec','update_test_status']   
#test_status_list=['CREATED','INQUIRED','PLANNED','TEST_OBJECTS_AVAILABLE','IN_PROGRESS','PERFORMED','ASSESSED','ASSESSED_REVIEWED','COMMENTED','COMPLETED','CHANGED','PAUSED']
test_status_list=['CREATED','INQUIRED','PLANNED','TEST_OBJECTS_AVAILABLE','IN_PROGRESS','PERFORMED','Query only']      
teamlist=['ETC31','ETC32','ETC33','ETC34','ETC35']                                       
#config参数
download_dir = ''# 下载地址
plan_file = ''  #测试计划表存放地址
test_dir = ''  # 测试文件夹根目录
raw_data_dir=''  #原始数据存放地址
approval_list = []  #审批人
conclusion_list = [] #结论
equipment_list = []  #设备清单
recipients_list=[]   #收件人
StringVarlist1 = []  # 存储后测试test id
StringVarlist2 = []  # 存储后测试test name
StringVarlist5 = []  # 存储文件夹名称
StringVarlist6=[]    #存储数据列特征
IntVarlist = []  # 存储单选按钮状态
first_approval_current_value = -1  # 存储审批人1默认值 
second_approval_current_value = -1  # 存储审批人2默认值
update_testinfo_current_value=0       #是否更新测试信息表
create_folder_current_value=0         #是否创建文件夹     
download_spec_current_value=0         #是否下载测试规范 
with_foldername_current_value=0       #是否文件夹附带名称
cobo_test_type_current_value = 0  # 存储待更新试验类型值
folder_list = []            #文件夹名称
sheet_name_current_value = ''  # 存储计划表分页名称
begin_row_current_value = ''  # 计划表数据开始行数
testinfo_download_current_value = ''  # 测试信息(下载表)
testinfo_plan_current_value = ''  # 测试信息(计划表)
object_plan_current_value = ''  # 样品信息(计划表)
addition_info_current_value =''   #附加信息
window_width=''             #窗口宽度
window_height=''            #窗口长度
search_test_criteria_current_value=''  #测试信息搜索字段
search_object_criteria_current_value=''  #样品信息搜索字段
search_test_criteria1_current_value=''  #测试信息搜索字段2
compress_picture_current_value=0  #压缩图片
time_sort_current_value=0  #时间排序
curve_generate_current_value=0  #生成曲线
hide_page_current_value=0  #隐藏浏览器
cobo_register_type_current_value=0 #登记类型
user_name='N.A'             #用户名
equipment_file_version=''   #设备信息表版本日期
data_character_current_value=''
sampling_frequency_current_value=''
column_list=[]
x_unit_current_value=''        #x轴单位
coordinate_typle_current_value=1        #坐标轴类型
hyperlink_column_current_value=1        #超链接列号
overwrite_current_value=1   #是否覆盖原数据
cobo_team_name_current_value='0'   #小组名

#网页地址
laboratories="https://rb-wam.bosch.com/edate/edate.html#laboratories;Tree:263326,Table:Overview:LaboratoryTestSheetView" #实验室主页
search="https://rb-wam.bosch.com/edate/edate.html#search"  #搜索页面
#页面元素xpath
search_load='/html/body/div[5]/div/div[2]/div/div/div[2]/div/div/div/div[1]/div/div/div[5]/div/div/div/img'
search_document='/html/body/div[9]/div//span[text()="Document"]'   
search_test='/html/body/div[9]/div//span[text()="Test"]'
search_test_object='/html/body/div[9]/div//span[text()="Test Object"]'
search_search_items='/html/body/div[5]/div/div[2]/div/div/div[2]/div/div/div/div[6]/div/div/div/div'
search_export='/html/body/div[5]/div/div[2]/div/div/div[4]/div/div[1]/div/div/div[7]/div/div'
search_excel='//*[@id="test-searchMenu-Export-Excel(Allrows)"]'
search_add='/html/body/div[5]/div/div[2]/div/div/div[2]/div/div/div/div[4]/div/div[2]/div[1]/div/div/div[1]/div/div'
search_add_test='/html/body/div[9]/div/div[7]/div/span'
search_add_test_id='/html/body/div[11]/div[2]/div[24]/div/span'
search_download_spec='/html/body/div[9]/div[2]/div[1]/div/div[2]/div/div[1]/div[12]/fieldset/div/div/div/div/div[1]/div/div[1]/div/div[3]/div/div[3]/div/div'
search_spec_close_window='/html/body/div[9]/div[1]/div/div[2]/table/tbody/tr/td[4]/div'
search_download_spec_save='/html/body/div[9]/div[2]/div[2]/div/div/div[1]/div/div'
search_download_spec_cancel='/html/body/div[9]/div[2]/div[2]/div/div/div[2]/div/div'
search_filter='//*[@id="test-Filter"]'
search_process='/html/body/div[5]/div/div[2]/div/div/div[4]/div/div[2]/div/div/div[2]/div[5]/div[2]'
export_process='/html/body/div[9]/div[2]/div[1]/div[3]'

laboratories_item='/html/body/div[5]/div/div[1]/div/div[1]/div[4]/div[2]/div[1]' 
laboratories_open='/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[5]/div/div/div/img'
laboratories_input='//*[@id="test-treeFilterInput-input"]'
laboratories_report='//*[@id="test-headerView-report"]'
laboratories_generate='/html/body/div[9]/div[2]/div[1]/div/div[2]/div[1]/div/div/div'
laboratories_result_files='/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[4]/div/div[1]/div[2]/div/div[2]/div/div'
laboratories_documents='/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[4]/div/div[1]/div[2]/div/div[3]/div/div'
laboratories_new='/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[4]/div/div[2]/div/div[1]/div/div/div[3]/div/div/div/img' 
laboratories_result_file='//*[@id="test-tableMenu-New-ResultFile"]'
laboratories_result_file_img="/html/body/div[9]/div/div/div/span/img"
laboratories_upload_document='/html/body/div[9]/div/div[1]/div/span'
laboratories_upload_document_img='/html/body/div[9]/div/div[1]/div/span/img'
#laboratories_add_document='//*[@id="test-dialogEdateDocument-files-button-Add"]' #这种地址定位不到
laboratories_add_document='/html/body/div[9]/div[2]/div[1]/div/div[1]/div/div[1]/div/div[2]/div/input[1]' #上传报告和结果都只能用这个
#laboratories_add_result='/html/body/div[9]/div[2]/div[1]/div/div[1]/div/div[1]/div/div[2]/div/input[2]' #不能用
laboratories_title='//*[@id="test-dialogEdateDocument-documentTitle-input"]'   
laboratories_comment='/html/body/div[9]/div[2]/div[1]/div/div[7]/div/div[1]/div/div/textarea'
                       

laboratories_save='/html/body/div[9]/div[2]/div[2]/div/div/div[1]/div/div'
laboratories_quickinfo='/html/body/div[5]/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[3]/fieldset/div/div/div[1]/div/div[2]/div'
alert_button='/html/body/div[9]/div[2]/div[2]/div/div/div/div/div'  #下载报告之后跳出的提示对话框里的ok按钮
laboratories_navigation_open='/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[5]/div/div/div'
laboratories_navigation_open_download='/html/body/div[9]/div[2]/div[1]/div/div[2]/div/div[1]/div[12]/fieldset/div/div/div/div/div[1]/div/div[1]/div/div[3]/div/div[3]/div'
laboratories_navigation_open_download_close='/html/body/div[9]/div[1]/div/div[2]/table/tbody/tr/td[4]/div'
laboratories_qucikinfo_test_status='/html/body/div[5]/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[3]/fieldset/div/div/div[3]/div/div[2]/div/div'
laboratories_overview_object='/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[4]/div/div[2]/div/div[3]/div/div[2]/div[4]/div/div[1]/div/div[2]/div/table/tbody[2]/tr[2]/td/div/table/tbody[2]/tr[1]'
laboratories_overview_object_ico='/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[4]/div/div[2]/div/div[3]/div/div[2]/div[4]/div/div[1]/div/div[2]/div/table/tbody[2]/tr[2]/td/div/table/tbody[2]/tr[1]/td[1]/div/div'
laboratories_qucikinfo_object_status='/html/body/div[5]/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[3]/fieldset/div/div/div[2]/div/div[2]/div'
laboratories_qucikinfo_object_deliverdate='/html/body/div[5]/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[5]/fieldset/div/div/div[3]/div/div[2]/div'
laboratories_navigation_testitem='/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[3]/div[2]/table/tbody/tr/td/div/div[2]/div/div[2]/div/div/span[2]'
laboratories_navigation_ico='/html/body/div[5]/div/div[2]/div/div/div/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[15]/div'
laboratories_navigation_ico_changestatus='/html/body/div[9]/div/div[2]/div'
laboratories_navigation_ico_changestatus_created='//*[@id="x-menu-el-test-treeMenu-ChangeStatus-▶CREATED"]'
laboratories_navigation_ico_changestatus_created_created='/html/body/div[9]/div[2]/div[2]/div/div/div[1]/div/div'
laboratories_navigation_ico_changestatus_inquired='//*[@id="x-menu-el-test-treeMenu-ChangeStatus-▶INQUIRED"]'
laboratories_navigation_ico_changestatus_inquired_inquired='/html/body/div[9]/div[2]/div[2]/div/div/div[1]/div/div'
laboratories_navigation_ico_changestatus_planned='//*[@id="x-menu-el-test-treeMenu-ChangeStatus-▶PLANNED"]'
laboratories_navigation_ico_changestatus_planned_planned='/html/body/div[9]/div[2]/div[2]/div/div/div[1]/div/div'
laboratories_navigation_ico_changestatus_testobjectsavailable='//*[@id="x-menu-el-test-treeMenu-ChangeStatus-▶TEST_OBJECTS_AVAILABLE"]'
laboratories_navigation_ico_changestatus_testobjectsavailable_ok='/html/body/div[9]/div[2]/div[2]/div/div/div[1]/div/div'
laboratories_navigation_ico_changestatus_inprogress='//*[@id="x-menu-el-test-treeMenu-ChangeStatus-▶IN_PROGRESS"]'
laboratories_navigation_ico_changestatus_inprogress_startdate='/html/body/div[9]/div[2]/div[1]/div/div[1]/div/div[1]/div/div/div/input[1]'
laboratories_navigation_ico_changestatus_inprogress_start='/html/body/div[9]/div[2]/div[2]/div/div/div[1]/div/div'
laboratories_navigation_ico_changestatus_inprogress_cancel='/html/body/div[9]/div[2]/div[2]/div/div/div[2]/div/div'
laboratories_navigation_ico_changestatus_performed='//*[@id="x-menu-el-test-treeMenu-ChangeStatus-▶PERFORMED"]'
laboratories_navigation_ico_changestatus_performed_startdate='/html/body/div[9]/div[2]/div[1]/div/div[1]/div/div[1]/div/div/div/input[1]'
laboratories_navigation_ico_changestatus_performed_enddate='/html/body/div[9]/div[2]/div[1]/div/div[2]/div/div[1]/div/div/div/input[1]'
laboratories_navigation_ico_changestatus_performed_complete='/html/body/div[9]/div[2]/div[2]/div/div/div[1]/div/div'
laboratories_navigation_ico_changestatus_assessed='//*[@id="x-menu-el-test-treeMenu-ChangeStatus-▶ASSESSED"]'
laboratories_navigation_ico_changestatus_assessedreviewed='//*[@id="x-menu-el-test-treeMenu-ChangeStatus-▶ASSESSED_REVIEWED"]'
                                                            

max_progress_value=100
current_progress=0

#各种地址
current_dir=os.getcwd()                                                      #程序文件地址
test_dir_unc=''                                                              #测试数据文件夹UNC地址
plan_file_unc=''                                                             #计划表UNC地址
blowser_path='C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'  #浏览器地址
driver_path = os.path.join(current_dir, 'msedgedriver.exe')                  #浏览器驱动地址
driver_package_path=os.path.join(current_dir, 'edgedriver_win64.zip')        #驱动安装包地址
driver_basic_url='https://msedgedriver.microsoft.com'                        #下载驱动网页地址
local_driver_path=r'\\bosch.com\dfsrb\DFSCN\Loc\CNG\ED\ENG\ETR-CN\01 Lab Management\1.9_Team_Management\02_ETC31\10_Environment\edataoper\bowser driver\msedgedriver.exe'  #本地浏览器驱动地址
#equipment_file_path=r'\\bosch.com\dfsrb\DFSCN\Loc\CNG\ED\ENG\ETR-CN\01 Lab Management\1.9_Team_Management\02_ETC31\10_Environment\edataoper\equipment_list'         #设备信息表文件地址
equipment_file_path=r'\\bosch.com\dfsrb\DfsCN\LOC\CNG\RBCC_Public\RBCC_Transfer\zhangshaoniu\equipment_list'
#equipment_file_path=r'\\bosch.com\dfsrb\DFSCN\Loc\CNG\ED\CTS-CN'
config_path=f'C:/Users/{register_name}/Desktop/conf.json'                              #配置文件地址
download_dir=f'C:/Users/{register_name}/Downloads'                                      #下载地址
print(config_path)


#加载配置文件
if (os.path.exists(config_path)):
    try:
        load_config(config_path)
    except Exception as e:
        tkinter.messagebox.showinfo('提示', '配置文件信息载入错误,请检查配置信息')
        print(str(e))
        #sys.exit(1)
        input("Press Enter to exit...")
else:
    tkinter.messagebox.showinfo('提示', '配置文件不存在')
    sys.exit(1)
    input("Press Enter to exit...")

formal_equipment_file_version=equipment_file_version
update_calibration_flag=None
#更新设备计量信息
try:
    update_calibration_flag=update_calibration_info()
except Exception as e:
    update_calibration_flag=False
    print(str(e))   
    input("Press Enter to exit...")

#生成程序主窗口    
try:
    window = Tk()
    window.protocol("WM_DELETE_WINDOW",close_window)  #将关闭窗口操作与close_window绑定
    w,h=get_screen_size()
    if window_width=='' or window_height=='':
        x=w//2-600                              #获取显示器屏幕中间的坐标,从而使得生成的窗口始终处于显示屏幕中间
        y=h//2-510
        window.geometry(f'1200x1020+{x}+{y}')
    else:
        x=w//2-int(window_width)//2
        y=h//2-int(window_height)//2
        window.geometry(f'{window_width}x{window_height}+{x}+{y}')
    window.title('edata自动化')

    menubar = tkinter.Menu(window)
    window.config(menu=menubar)
    
    #创建配置菜单
    config_menu = tkinter.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="配置", menu=config_menu)
    
    #添加一级菜单项
    config_menu.add_command(label="载入配置",command=select_config)
    config_menu.add_command(label="保存配置",command=lambda:save_config(config_path))
    config_menu.add_separator()


    #创建地址选择菜单
    address_menu = tkinter.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="地址", menu=address_menu)
    
    #添加一级菜单项
    address_menu.add_command(label="测试计划表",command=select_plan_file)
    address_menu.add_command(label="数据文件夹根目录",command=select_test_dir)
    address_menu.add_command(label="原始数据文件夹根目录",command=select_raw_data_dir)
    address_menu.add_command(label="查看地址",command=view_address)
    address_menu.add_separator()
    
    #创建设备菜单
    equipment_menu = tkinter.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="设备", menu=equipment_menu)
    
    #添加一级菜单项
    equipment_menu.add_command(label="新增设备",command=lambda:add_equipment(window))
    equipment_menu.add_command(label="删除设备",command=lambda:select_delete(window,'删除设备',equipment_list))
    equipment_menu.add_command(label="查看设备",command=lambda:view_data(equipment_list,'查看设备信息'))
    equipment_menu.add_separator()
    
    #创建审批人菜单
    approval_menu = tkinter.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="审批人", menu=approval_menu)
    
    # 添加一级菜单项
    approval_menu.add_command(label="新增审批人",command=lambda:add_approval(window))
    approval_menu.add_command(label="删除审批人",command=lambda:select_delete(window,'删除审批人',approval_list))
    approval_menu.add_command(label="查看审批人",command=lambda:view_data(approval_list,'查看审批人信息'))
    approval_menu.add_separator()

    #创建结论菜单
    conclusion_menu = tkinter.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="结论", menu=conclusion_menu)
    
    # 添加一级菜单项
    conclusion_menu.add_command(label="新增结论",command=lambda:add_conclusion(window,'button38'))
    conclusion_menu.add_command(label="删除结论",command=lambda:select_delete(window,'删除结论',conclusion_list))
    conclusion_menu.add_command(label="查看结论",command=lambda:view_data(conclusion_list,'查看测试结论信息'))
    conclusion_menu.add_separator()

    #创建邮件收件人菜单
    recipients_menu = tkinter.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="收件人", menu=recipients_menu)
    
    # 添加一级菜单项
    recipients_menu.add_command(label="新增收件人",command=lambda:add_recipient(window))
    recipients_menu.add_command(label="删除收件人",command=lambda:select_delete(window,'删除收件人',recipients_list))
    recipients_menu.add_command(label="查看收件人",command=lambda:view_data(recipients_list,'查看收件人'))
    recipients_menu.add_separator()

    #生成tab栏
    style = ttk.Style()
    # 设置Notebook样式
    style.configure("Custom.TNotebook", background="#f0f0f0", borderwidth=0)
    style.configure("Custom.TNotebook.Tab", 
                    padding=[20, 5],  # 增加内边距
                    font=("Microsoft YaHei", 10, "bold"),
                    background="#e0e0e0",
                    foreground="#333333")
    style.map("Custom.TNotebook.Tab",
              background=[("selected", "#1a5fb4"),   # 更深的蓝色
                      ("active", "#d0d0d0")],    # 悬停状态
              foreground=[("selected", "black"),     # 选中时白色文字
                      ("active", "#111111")],    # 悬停时深色文字
              relief=[("selected", "sunken"),       # 选中时凹陷效果
                  ("!selected", "ridge")])      # 未选中时凸起效果)

    # 创建Notebook
    tabControl = ttk.Notebook(window, width=1200, height=1100, style="Custom.TNotebook")

    # 创建Tab页
    tab1 = ttk.Frame(tabControl)
    tab2 = ttk.Frame(tabControl)
    tab3 = ttk.Frame(tabControl)
    tab4 = ttk.Frame(tabControl)

    # 添加Tab页
    tabControl.add(tab1, text='📄 报告生成')
    tabControl.add(tab2, text='📊 计划表更新')
    tabControl.add(tab3, text='📈 曲线生成')
    tabControl.add(tab4, text='📁 数据拷贝')

    # 布局
    #tabControl.pack(expand=True, fill="both", padx=10, pady=10)

    tabControl.grid(padx=10, pady=10)
    generate_report = GenerateTestReport(tab1)
    print('generate_report window is ok ')
    update_test_plan = UpdateTestPlan(tab2)   
    print('upldate_test_plan window is ok ')
    test_cuve=TestCurve(tab3)
    print('test_curve window is ok ')
    data_copy = AdvancedCopyApp(tab4)
    print('data_copy window is ok ')

    #显示登陆账号用户名
    showinfo(generate_report.report_status,f'登录账号用户名:{register_name}','black')
    plan_file_drive_letter=None
    test_dir_driver_letter=None
    plan_file_drive_letter_unc_path=None
    test_dir_driver_letter_unc_path=None
    #获取相应盘符的unc地址    
    try:
        plan_file_drive_letter=plan_file[0]  #获取计划表地址的盘符
    except Exception as e:
        showinfo(generate_report.report_status,'未设置计划表地址','red')
    try:
        test_dir_driver_letter=test_dir[0]   #获取测试文件夹根目录盘符
    except Exception as e:
        showinfo(generate_report.report_status,'未设置数据文件夹地址','red')    
  
    if plan_file_drive_letter is not None:
        showinfo(generate_report.report_status,f'计划表盘符:{plan_file_drive_letter}','black')
        plan_file_drive_letter_unc_path=get_unc_path(plan_file_drive_letter)      #计划表地址盘符的unc地址
    if test_dir_driver_letter is not None:
        showinfo(generate_report.report_status,f'数据文件夹盘符:{test_dir_driver_letter}','black')
        test_dir_driver_letter_unc_path=get_unc_path(test_dir_driver_letter)      #数据文件夹地址盘符的unc地址
    if plan_file_drive_letter_unc_path is not None:
        plan_file_unc=plan_file_drive_letter_unc_path+'//'+plan_file.split(':',1)[1]         #计划表unc地址
        showinfo(generate_report.report_status,f'计划表盘符unc地址:{plan_file_drive_letter_unc_path}','black')
    else:
        plan_file_unc=plan_file
        showinfo(generate_report.report_status,f'计划表地址:{plan_file}','black')

    if test_dir_driver_letter_unc_path is not None:
        test_dir_unc=test_dir_driver_letter_unc_path+'//'+test_dir.split(':',1)[1]           #数据文件夹unc地址
        showinfo(generate_report.report_status,f'数据文件夹盘符unc地址:{test_dir_driver_letter_unc_path}','black')
    else:    
        test_dir_unc=test_dir
        showinfo(generate_report.report_status,f'数据文件夹地址:{test_dir}','black')

    #获取edge浏览器的版本号和浏览器驱动文件的版本号
    blowser_version=get_file_version(blowser_path)
    driver_version=get_file_version(driver_path)
    #edge_version=powershell_cmd(f'(get-Item "C:\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe").VersionInfo.ProductVersion') #也可以用执行powershell命令方式获取版本号
    showinfo(generate_report.report_status,f'浏览器版本号:{blowser_version}','black')
    showinfo(generate_report.report_status,f'驱动版本号:{driver_version}','black')
    showinfo(generate_report.report_status,f'equipment_file_version:{equipment_file_version}','black')
    #提示equipment_file_version更新信息
    if update_calibration_flag==True:
        showinfo(generate_report.report_status,'设备计量信息发生了更新变动,请在看见此提示时点击保存配置','red')
        showinfo(generate_report.report_status,f'update the equipment_file_version from {formal_equipment_file_version}------>{equipment_file_version}','red')
    elif update_calibration_flag==False:    
        showinfo(generate_report.report_status,'设备计量信息更新失败','red')
    window.mainloop()
    #放在window.mainloop()语句之后不会执行，因为mainloop()执行后在这一步一直等待
except Exception as e:
    print(str(e))
    input("Press Enter to exit...")   


