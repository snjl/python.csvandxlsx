from openpyxl import Workbook
import csv


def save_excel(file_name, headers, rows):
    """
    将表头和数据写入新生成的excel文件中，生成*.xlsx文件
    :param file_name:文件名
    :param headers:文件第一行的表头元素，list存储
    :param rows:文件数据，每一行为一个list
    """
    # 在内存中创建一个workbook对象，而且会至少创建一个 worksheet
    wb = Workbook()
    # 获取当前活跃的worksheet,默认就是第一个worksheet
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    # 保存格式必须以.xlsx结尾，不然加上.xlsx
    if file_name.endswith(".xlsx") is not True:
        file_name += '.xlsx'
    wb.save(filename=file_name)
    wb.close()


def save_csv(file_name, headers, rows):
    """
    将表头和数据写入新生成的文件中
    :param file_name: 生成文件名，例如xxx.csv
    :param headers:文件第一行的表头元素，list存储
    :param rows:文件数据，每一行为一个list
    """
    # 保存格式必须以.csv，不然加上.csv
    if file_name.endwith(".csv") is not True:
        file_name += '.csv'
    with open(file_name, 'a', encoding='utf8', errors='ignore', newline='') as f:
        f_csv = csv.writer(f)
        f_csv.writerow(headers)
        f_csv.writerows(rows)
