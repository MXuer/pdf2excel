import pdfplumber
import pandas as pd
import json
import re
from collections import defaultdict
# 打开PDF文件
pdf_path = "data/20241101011-FLPJade-pump泵体6491902.pdf"
output_path = "extracted_data.xlsx"

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
    # 使用 pdfplumber 提取PDF内容
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
                    # print(text)
                    continue
                if output_dict.get(type_, None) is None:
                    # print('-'*30)
                    # print(f'{type_}|{contents}')
                    # print('-'*30)
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
                            output_dict[article_number_range] = each[2]
                    table_b = tables[-1]
                    output_dict['金属颗粒长度'] = table_b[1][4]
                    output_dict['金属颗粒宽度'] = table_b[1][5]
                    output_dict['非金属颗粒长度'] = table_b[2][4]
                    output_dict['非金属颗粒宽度'] = table_b[2][5]
    return output_dict

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

from glob import glob
pdf_files = glob(f'data/*.pdf')

for pdf_path in pdf_files:
    output_dict = read_one_pdf(pdf_path)
    for k, v in output_dict.items():
        final_results[k].append(v)
# 整理并导出到Excel
sample_df = pd.DataFrame(final_results)
sample_df.to_excel(output_path, index=False)
