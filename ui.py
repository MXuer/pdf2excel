import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from glob import glob

def read_one_pdf(pdf_path):
    output_dict = {
        '样品型号': '',
        '操作员': '',
        '取样于': '',
        '研究日期': '',
        '零件编号': '',
        '表面积': '',
        '取样数量': '',
        '检测方法/清洁度等级': '',
        '金属颗粒长度': '',
        '金属颗粒宽度': '',
        '非金属颗粒长度': '',
        '非金属颗粒宽度': '',
    }
    type2type = {
        '产品面积': '表面积'
    }
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            page_text = page_text.split('\n')
            for text in page_text:
                if ':' in text:
                    type_, contents = text.split(':', 1)
                    type_, contents = type_.strip(), contents.strip()
                    type_ = type2type.get(type_, type_)
                else:
                    continue
                if output_dict.get(type_, None) is None:
                    continue
                if type_ == '表面积':
                    contents = int(float(contents.split('cm²')[0]) * 100)
                output_dict[type_] = contents
            if 'Particles table' in page_text:
                tables = page.extract_tables()
                if tables:
                    table_a = tables[0]
                    for each in table_a:
                        article_number_range = re.sub("\(.*?\)", '', each[0]).strip() + ' 颗粒数'
                        if re.findall('µm', article_number_range):
                            article_number_range = article_number_range.replace('µm -', '-')
                            if "=" not in article_number_range:
                                article_number_range = article_number_range.replace('>', '>=')
                            output_dict[article_number_range] = each[2]
                    table_b = tables[-1]
                    output_dict['金属颗粒长度'] = table_b[1][4]
                    output_dict['金属颗粒宽度'] = table_b[1][5]
                    output_dict['非金属颗粒长度'] = table_b[2][4]
                    output_dict['非金属颗粒宽度'] = table_b[2][5]
    return output_dict

def process_pdfs():
    pdf_files = glob(f'{input_folder.get()}/*.pdf')
    final_results = {
        "样品型号": [],
        "操作员": [],
        "取样于": [],
        "研究日期": [],
        "零件编号": [],
        "表面积": [],
        "取样数量": [],
        "检测方法/清洁度等级": [],
        "金属颗粒长度": [],
        "金属颗粒宽度": [],
        "非金属颗粒长度": [],
        "非金属颗粒宽度": [],
        "6 - 15 µm 颗粒数": [],
        "15 - 25 µm 颗粒数": [],
        "25 - 50 µm 颗粒数": [],
        "50 - 100 µm 颗粒数": [],
        "100 - 150 µm 颗粒数": [],
        "150 - 200 µm 颗粒数": [],
        "200 - 400 µm 颗粒数": [],
        "400 - 600 µm 颗粒数": [],
        "600 - 1000 µm 颗粒数": [],
        ">= 1000 µm 颗粒数": []
    }
    for pdf_path in pdf_files:
        output_dict = read_one_pdf(pdf_path)
        for k, v in output_dict.items():
            final_results[k].append(v)
    sample_df = pd.DataFrame(final_results)
    sample_df.to_excel(output_path.get(), index=False)
    messagebox.showinfo("完成", "数据已成功提取并保存到Excel文件中。")

def select_input_folder():
    folder_selected = filedialog.askdirectory()
    input_folder.set(folder_selected)

def select_output_file():
    file_selected = filedialog.asksaveasfilename(defaultextension=".xlsx")
    output_path.set(file_selected)

# 创建主窗口
root = tk.Tk()
root.title("PDF to Excel Converter")

# 输入文件夹选择
tk.Label(root, text="选择PDF文件夹:").grid(row=0, column=0, padx=10, pady=10)
input_folder = tk.StringVar()
tk.Entry(root, textvariable=input_folder, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="浏览", command=select_input_folder).grid(row=0, column=2, padx=10, pady=10)

# 输出文件选择
tk.Label(root, text="选择输出Excel文件:").grid(row=1, column=0, padx=10, pady=10)
output_path = tk.StringVar()
tk.Entry(root, textvariable=output_path, width=50).grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="浏览", command=select_output_file).grid(row=1, column=2, padx=10, pady=10)

# 处理按钮
tk.Button(root, text="开始处理", command=process_pdfs).grid(row=2, column=0, columnspan=3, pady=20)

# 运行主循环
root.mainloop()