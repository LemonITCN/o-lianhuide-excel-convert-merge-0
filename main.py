# coding=utf-8
import tkinter
import tkinter.messagebox
from tkinter.filedialog import (askopenfilename,
                                asksaveasfilename)
import xlrd
import xlwt
from xlwt import easyxf, Workbook

window = tkinter.Tk()
global excel_1_btn
global excel_1_label
global excel_2_btn
global excel_2_label
global generate_btn

global excel_1_path
global excel_2_path
global generate_path


def create_ui():
    # 创建窗口
    global window, excel_1_btn, excel_1_label, excel_2_btn, excel_2_label, generate_btn
    window.geometry('480x320+200+200')
    window.resizable(0, 0)
    window.title("LHD")

    # 创建按钮和标签控件
    excel_1_btn = tkinter.Button(window, text="选择表格1", command=btn_action_excel_1)
    excel_1_btn.place(x=20, y=20)

    excel_1_label = tkinter.Label(window, text="您还没有选择表格1文件")
    excel_1_label.place(x=20, y=50)

    excel_2_btn = tkinter.Button(window, text="选择表格2", command=btn_action_excel_2)
    excel_2_btn.place(x=20, y=100)

    excel_2_label = tkinter.Label(window, text="您还没有选择表格2文件")
    excel_2_label.place(x=20, y=130)

    generate_btn = tkinter.Button(window, text="生成表格3", command=btn_action_generate)
    generate_btn.place(x=20, y=180)

    # 开始
    window.mainloop()


def btn_action_excel_1():
    global excel_1_path, excel_1_label
    excel_1_path = askopenfilename()
    excel_1_label.config(text="已选择：" + excel_1_path)


def btn_action_excel_2():
    global excel_2_path, excel_2_label
    excel_2_path = askopenfilename(title='请选择Excel表格1文件', filetypes=[('xls', '*.xls'), ('xlsx', '*.xlsx')])
    excel_2_label.config(text="已选择：" + excel_2_path)


def btn_action_generate():
    global generate_path
    generate_path = asksaveasfilename()
    generate_output_file()
    tkinter.messagebox.showinfo('提示', '生成完毕，已将生成的文件保存到：' + generate_path)


def generate_output_file():
    # 创建新的Excel
    excel_out_book = xlwt.Workbook(encoding="utf-8")
    excel_out_sheet = excel_out_book.add_sheet('Sheet1')
    # 处理数据
    excel_1 = xlrd.open_workbook(excel_1_path)
    excel_1_sheet = excel_1.sheet_by_index(0)
    # 计算并存储表1数据
    excel_1_data = {}
    for excel_1_row in excel_1_sheet.get_rows():
        excel_1_data[str(excel_1_row[0].value)] = excel_1_row
    # 根据表二数据开始生成表三数据
    excel_2 = xlrd.open_workbook(excel_2_path)
    excel_2_sheet = excel_2.sheet_by_index(0)
    excel_out_sheet.write(0, 0, 'sku')
    excel_out_sheet.write(0, 1, 'item_id')
    excel_out_sheet.write(0, 2, 'ebay中listing访问地址')
    excel_out_sheet.write(0, 3, '价格(USD)')
    excel_out_sheet.write(0, 4, '上架时间')
    excel_out_sheet.write(0, 5, 'Sold_qty')
    excel_out_sheet.write(0, 6, 'Sold_for')
    excel_out_sheet.write(0, 7, 'Ad_fees')
    excel_out_sheet.write(0, 8, '广告费率')
    excel_out_sheet.write(0, 9, '销售额')
    excel_out_sheet.write(0, 10, '售出')
    excel_out_sheet.write(0, 11, '广告投入率')
    row_index = 0
    style_percent = easyxf(num_format_str='0.00%')
    for excel_2_row in excel_2_sheet.get_rows():
        if row_index > 0:
            row_id = str(excel_2_row[0].value)
            # A列是从表2AB列
            excel_out_sheet.write(row_index, 0, excel_2_row[27].value)
            # B列是从表1A列
            excel_out_sheet.write(row_index, 1, excel_2_row[0].value)
            # C列是从表2AH列
            excel_out_sheet.write(row_index, 2, excel_2_row[33].value)
            # D列是从表2AD列
            excel_out_sheet.write(row_index, 3, excel_2_row[29].value)
            # E列是从表2AC列
            excel_out_sheet.write(row_index, 4, excel_2_row[28].value)
            # J列是从表2J列
            excel_out_sheet.write(row_index, 9, excel_2_row[9].value)
            # K列是 从表2 O列
            excel_out_sheet.write(row_index, 10, excel_2_row[15].value)
            if row_id in excel_1_data:
                # F列是从表1B列
                excel_out_sheet.write(row_index, 5, excel_1_data[row_id][1].value)
                # G列是从表1C列
                excel_out_sheet.write(row_index, 6, excel_1_data[row_id][2].value)
                # H列是从表1D列
                excel_out_sheet.write(row_index, 7, excel_1_data[row_id][3].value)
                # I列是计算得出的，用表3H列 / 表3G列
                excel_out_sheet.write(row_index, 8, (excel_1_data[row_id][3].value / excel_1_data[row_id][2].value),
                                      style_percent)
                # L列是计算得出的，用表3H列 / 表3J列
                excel_out_sheet.write(row_index, 11, (excel_1_data[row_id][3].value / float(excel_2_row[9].value[1:])),
                                      style_percent)
        row_index = row_index + 1
    excel_out_book.save(generate_path)


if __name__ == '__main__':
    create_ui()

# 注意需要修改xlrd源码
# Python\lib\site-packages\xlrd\compdoc.py
# 大约426行左右，注释掉下面这行
# raise CompDocError("%s corruption: seen[%d] == %d" % (qname, s, self.seen[s]))