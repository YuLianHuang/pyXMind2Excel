import json
import xmind
import jmespath
# import openpyxl
import xlwt
import time
import logging
import sys
from pathlib import Path


def find_items(data):
    # try:
    results = []

    if isinstance(data, dict):
        title = data.get('title')
        if title is not None and isinstance(title, str) and 'TC' in title:
            results.append(data)
        elif 'topics' in data:
            results.extend(find_items(data.get('topics')))
        else:
            pass
    elif isinstance(data, list):
        for item in data:
            results.extend(find_items(item))
    return results
    # except Exception as e:
    #     logging.error(f"Error occurred while finding items: {e}")
    #     raise

def parse_data(data,autor,req_id):
    try:
        # print(data)
        topic_list = data[0]['topic']['topics']

        TC_data = find_items(topic_list)

        # TC_data = jmespath.search("topics[].topics[].topics[?contains(title,'TC')][]", target_data)
        TC_title = jmespath.search("[].title", TC_data)
        TC_title = [i.replace('：',':').split('TC:')[1] if i != '' else i for i in TC_title]
        TC_condition = jmespath.search("[].topics[].title[]", TC_data)
        TC_step = jmespath.search("[].topics[].topics[].title[]", TC_data)
        TC_expect = jmespath.search("[].topics[].topics[].topics[].title[]", TC_data)

        len_diff_cond = len(TC_title) - len(TC_condition)
        TC_condition.extend([''] * len_diff_cond)
        len_diff_step = len(TC_title) - len(TC_step)
        TC_step.extend([''] * len_diff_step)
        len_diff_expect = len(TC_title) - len(TC_expect)
        TC_expect.extend([''] * len_diff_expect)
        
        # TC_condition = [i.replace('：',':').split('FC:')[1] if i != '' else i for i in TC_condition]
        # TC_step = [i.replace('：',':').split('ST:')[1] if i != '' else i for i in TC_step]
        # TC_expect = [i.replace('：',':').split('ER:')[1] if i != '' else i for i in TC_expect]

        cases = zip(TC_title,TC_condition,TC_step,TC_expect)
        
        TC_list = []
        for case in cases:
            case = list(case)
            case.insert(0,'')
            case.insert(2,'非自动化用例')
            case.insert(3,f'{req_id}')
            case.insert(4,f'{autor}')
            case.insert(5,'步骤一')
            case.insert(9,'')
            TC_list.append(case)
        return TC_list
    
    except Exception as e:
        logging.error(f"Failed to parse data: {e}")
        raise


def write_excel(data):
    try:
        if not isinstance(data,(list,tuple)):
            raise ValueError("data must be a list or tuple")
        
        wb = xlwt.Workbook()
        sheet1 = wb.add_sheet('用例模板')
        ws = wb.get_sheet(0)

        title = ['目录层级','*用例名称','用例类型','关联需求编号','作者','步骤名','前置条件','执行步骤','预期结果','描述']
        for count,tt in enumerate(title):
            ws.write(0,count,tt)
        for row_index,row_data in enumerate(data):
            if not isinstance(row_data,(list,tuple)):
                    raise ValueError("Each row of data must be a list or tuple")
            for col_index,col_data in enumerate(row_data):
                ws.write(row_index+1,col_index,col_data)

        now = time.strftime(r'%Y%m%d_%H%M%S',time.localtime())
        filename = f'testcase_{now}.xls'
        wb.save(filename)
        logging.info(f"Excel file saved successfully: {filename}")
        return filename

    except Exception as e:
        logging.error(f"Failed to save Excel file: {e}")
        raise


def run(xmind_filepath,autor="",req_id=""):
    xmind_filepath = xmind_filepath.strip('"')
    if not Path(xmind_filepath):
        print("文件不存在，请重新输入")
        return
    if not xmind_filepath.endswith('.xmind'):
        print("文件格式不正确，请重新输入")
        return
    Workbook = xmind.load(xmind_filepath)
    data = Workbook.getData()
    filename = write_excel(parse_data(data,autor,req_id))
    print(f"导出excel成功！文件名为：{filename}")


def main():
    try:
        filepath = input("请输入xmind文件绝对路径（也可以直接将文件拖动到当前窗口）：\n")
        autor = input("请输入用例作者姓名：\n")
        req_id = input("请输入关联需求编号：\n")
        run(filepath,autor,req_id)
    except Exception as e:
        print(f"发生错误：{e}")
        raise e
    finally:
        input("\n\nPress Enter to exit...")
    # sys.argv[0] 是脚本名称，sys.argv[1] 是第一个参数，依此类推
    # if len(sys.argv) < 2:
    #     sys.exit(1)

    # arg1 = sys.argv[1]
    

if __name__ == "__main__":
    main()
    # run(r'E:\my_data\pppp\KSMP\2.DevelopmentDoc\2.5Test\2.5.5SystemTest\策略执行\测试案例\测试要点\KSMP_SES_1.1.6.0\KSMP-74 日终终止订单消息推送.xmind','黄雨炼','KSMP-74')
