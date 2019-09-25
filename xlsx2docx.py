# -*-coding: utf-8-*-
# editor: MikeChan
# email: m7807031@gmail.com
# license: BSD-3

# OS import
import os
import random
from datetime import datetime, date
import time
import re

# Pandas import
import pandas as pd

# Office-docx lib import
import docx
from docx.shared import Cm, Pt  #add unit for word
from docx.enum.text import WD_ALIGN_PARAGRAPH # deal with alignment
from docx.enum.table import WD_TABLE_ALIGNMENT
#from docx.shared import RGBColor

# ===== TODO ===== #
# 1. can't open large row excel
# 2. deal with folder's different name

# ===== Global Variables ===== #
spacing_folder = ""
input_folder = "input/"
output_folder = "output/"

# ===== Global Function ===== #
def timeIt(func):
    def wrapper(*args, **kw):
        t1 = time.time()
        res = func(*args, **kw)
        t2 = time.time()
        print("** timeIt Report: {} took {:.2f}s".format(func.__name__, t2-t1))
    return wrapper

# ====== Functions ===== #
def output_folder_dir_list(input_folder_path):
    dir_list = os.listdir(input_folder_path)
    return dir_list


def xlsx2df(xlsx_path):
    folder_number = xlsx_path.split("/")[-1]
    raw_df = pd.read_excel(xlsx_path + "/"+ u"台電跑報表_缺陷匯總表.xlsx")
    df = raw_df.loc[1::]
    df.columns = [u"序號",u"區間",u"距小號塔距離(m)",u"地物危險點座標", 
                  u"缺陷類型",u"缺陷級別",
                  u"實測平距",u"實測垂距",u"實測直線距離",
                  u"規範平距",u"規範垂距",u"規範直線距離",u"圖示"]
    
    df[u"區間"] = folder_number
    #print(df.head(2))
    print(u"-> {}台電跑報表_缺陷匯總表.xlsx read.".format(xlsx_path) )
    return df

@timeIt
def df2docx(df, output_folder, spacing_folder):
    if not os.path.exists(output_folder+spacing_folder):
        os.makedirs(output_folder+spacing_folder)
    
    docx_name = spacing_folder + u"缺陷匯總表.docx"

    # get basic info from df
    #row_length = df.shape[0]
    row_length = df.shape[0]
    col_length = df.shape[1]

    print("-> Total Row: {}".format(row_length))
    

    # open temp docx file
    doc = docx.Document()

    # adjust border
    section_0 = doc.sections[0]
    section_0.left_margin = Cm(1.27)
    section_0.right_margin = Cm(1.27)
    section_0.top_margin = Cm(1.27)
    section_0.bottom_margin = Cm(1.27)


    table_0 = doc.add_table(rows=2 ,cols=col_length, style='Table Grid')
    table_0.alignment = WD_TABLE_ALIGNMENT.CENTER
    # deal with heading
    table_0.cell(0,0).merge(table_0.cell(1,0))
    table_0.cell(0,1).merge(table_0.cell(1,1))
    table_0.cell(0,2).merge(table_0.cell(1,2))
    table_0.cell(0,3).merge(table_0.cell(1,3))
    table_0.cell(0,4).merge(table_0.cell(1,4))
    table_0.cell(0,5).merge(table_0.cell(1,5))
    table_0.cell(0,6).merge(table_0.cell(0,8))
    table_0.cell(0,9).merge(table_0.cell(0,11))
    table_0.cell(0,12).merge(table_0.cell(1,12))

    table_0.cell(0,0).text = u"序號"
    table_0.cell(0,1).text = u"區間"
    table_0.cell(0,2).text = u"距小號塔距離(m)"
    table_0.cell(0,3).text = u"地物危險點座標"
    table_0.cell(0,4).text = u"缺陷類型"
    table_0.cell(0,5).text = u"缺陷級別"
    table_0.cell(0,6).text = u"實測距離(m)"
    table_0.cell(0,9).text = u"規範要求安全距離(m)"
    table_0.cell(0,12).text = u"圖示"

    table_0.cell(1,6).text = u"平距"
    table_0.cell(1,7).text = u"垂距"
    table_0.cell(1,8).text = u"直線距離"
    table_0.cell(1,9).text = u"平距"
    table_0.cell(1,10).text = u"垂距"
    table_0.cell(1,11).text = u"直線距離"
    
    print("---> table header ok.")

    
    table_1 = doc.add_table(rows=row_length ,cols=col_length, style='Table Grid')
    table_cells_1 = table_1._cells

    table_1.alignment = WD_TABLE_ALIGNMENT.CENTER
    print("---> table added.")


    for i in range(row_length):
        print("-> Row Processing: {} / {}". format(i+1, row_length))
        row_cells = table_cells_1[i*col_length : (i+1)*col_length]
        for j in range(len(row_cells)):
            insert_value = df.iloc[i, j]
            if j == 0:
                insert_value = str(int(insert_value))
            if j>=6 and j <=11:
                insert_value = "{:.3f}".format(insert_value)
            if str(insert_value) == "nan":
                insert_value = ""
            row_cells[j].text = str(insert_value)

        #Add text to row_cells



    '''
    # fill the cell with df data
    for row_index in range(row_length):
        print("-> Row Processing: {} / {}". format(row_index+1, row_length))
        for col_index in range(col_length):
            insert_value = df.iloc[row_index, col_index]
            if col_index == 0:
                insert_value = str(int(insert_value))
            if col_index >= 6 and col_index <= 11:
                insert_value = "{:.3f}".format(insert_value)
            if str(insert_value) == "nan":
                insert_value = ""

            table_1.cell(row_index, col_index).text = str(insert_value)
    '''

    doc.save(output_folder + spacing_folder + "/" + docx_name)
    print("---> docx saved.")

# ====== main() ===== #
def main():
    dir_list = os.listdir(input_folder)
    for tar_folder in dir_list:
        print("*** {} is processing...".format( input_folder + tar_folder + "/"))
        single_df = xlsx2df(input_folder + "/" + tar_folder)
        df2docx(single_df, output_folder, tar_folder)



if __name__=="__main__":
    print("** xlsx2docx convertion start. **")

    main()

    print("** xlsx2docx convertion Done. **")