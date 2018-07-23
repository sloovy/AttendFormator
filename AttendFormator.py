import xlrd
import xlwt
from datetime import datetime
from datetime import time
import calendar
import argparse


'''
根据公司考勤表格式，从默认输出格式的xls转为考勤指定格式。

要点：
1.按照指定列格式输出；
2.两条上下班签到/签退记录合并为一条考勤记录；
3.填充没有考勤记录的空白天数；
4.注意个别的通宵加班情况，签退时间可能在第二天凌晨。

内部存储考勤数据表结构：
1.总表为dict，键是登记号码pin，键值是本月考勤记录表
2.考勤记录表为list，item内容是InputRecord列表。导入本pin全部记录后，按时间排序
3.输出时，重新按指定格式处理内容后输出。
4.考勤月份、工作表名、输出文件名，都以第一条有效记录的时间月份为准
'''
silent_mode = False
data_dic = {}   # 考勤数据总表dict
WORK_OFF_TIME = time(hour=18, minute=30)     # 下班时间定义 18:30
WHOLE_NIGHT_WORK_TIME = time(hour=7,minute=0)   # 7点前打卡视为通宵加班
XLS_BLANK_FILL_WEEKEND = "green"    # 输出表格中周末空白填充颜色
XLS_BLANK_FILL_EXCEPT = "red"

class InputRecord:
    '''读取输入表格转为内部记录对象'''
    #-- 输入表格列名枚举 --
    #dept_id #部门
    #name #姓名
    #pin #登记号码
    #check_time #日期时间
    #sensor_id #机器号
    #ssn #编号
    #verify_code #比对方式
    #card_no #卡号
    DEPT_ID = 0
    NAME = 1
    PIN = 2
    CHECK_TIME = 3

    def __init__(self, name, pin, check_time):
        self.name = name
        self.pin = pin
        self.check_time = check_time
    def __str__(self):
        return "name:{0}, pin:{1}, check_time:{2}".format(self.name, self.pin, self.check_time)

    #end class InputRecord
class OutputRecord:
    '''输出xls表格格式字段'''
    #pin        #登记号码
    #weekday    #星期几
    #name       #姓名
    #check_date #日期
    #check_in   #签到时间
    #check_out  #签退时间
    #except_code #异常类型，详见下列表
    PIN = 0
    WEEKDAY = 1
    NAME = 2
    CHECK_DATE = 3
    CHECK_IN = 4
    CHECK_OUT = 5
    EXCEPT_CODE = 6

    COLUMN_NAME = ["登记号码","星期几","姓名","日期","签到时间","签退时间","异常类型"]
    EXCEPT_CODE_TEXT = ["","上班未签到","下班未签退","通宵加班","上月末通宵加班"]
    COLUMN_COUNT = len(COLUMN_NAME)

    def __init__(self, pin, weekday, name, check_date, check_in=None, check_out=None, except_code=0):
        self.pin = pin
        self.weekday = weekday
        self.name = name
        self.check_date = check_date
        self.check_in = check_in
        self.check_out = check_out
        self.except_code = except_code
    def __str__(self):
        return "pin:{0}, weekday:{1}, name:{2}, date:{3}, checkin:{4}, checkout:{5}, except_code:{6}".format(
                    self.pin, self.weekday, self.name,
                    self.check_date, self.check_in, self.check_out,
                    self.except_code)

    #end class OutputRecord

def debug_print(output):
    if not silent_mode:
        print(output)
    pass
    #end debug_print()

def read_input_xls(input_file):
    '''读取输入表格到内部数据列表'''
    book = xlrd.open_workbook(input_file, encoding_override="gbk")
    if book.nsheets == 0:
        print("[ERR]No sheet find!")
        return 0

    sheet = book.sheet_by_index(0)
    debug_print("Sheet name:{0}, Columns:{1}, Rows:{2}".format(sheet.name, sheet.ncols, sheet.nrows))
    # 无数据则直接返回0
    if sheet.nrows <= 1:
        print("[ERR]No data!")
        return 0


    #DEPT_ID = 0
    #NAME = 1
    #PIN = 2
    #CHECK_TIME = 3
    # 第一行是列名标题
    record_title = InputRecord(sheet.cell_value(0,InputRecord.NAME),
                               sheet.cell_value(0,InputRecord.PIN),
                               sheet.cell_value(0,InputRecord.CHECK_TIME))
    debug_print(record_title)

    # 逐行读取数据，加入列表
    last_pin = 0     # 当前处理的登记号码。如果不同了要开新list
    record_list = []
    for rx in range(1, sheet.nrows):
        debug_print(sheet.row_values(rx))
        # 简单验证数据有效性: 1.pin不为0，2.名字长度至少2字符
        pin = int(sheet.cell_value(rx, InputRecord.PIN))
        if pin == 0:
            continue
        name = str(sheet.cell_value(rx,InputRecord.NAME)).strip()
        if len(name) <= 1:
            continue
        # 加入列表并记录pin
        checktime = get_input_checktime(sheet.cell_value(rx,InputRecord.CHECK_TIME))
        record = InputRecord(name, pin, checktime)

        #如果和上一条不同pin，则保存目前记录列表到总表后，开新list
        if pin != last_pin and last_pin != 0:
            data_dic[last_pin] = record_list.copy()     #直接赋值是引用地址！
            record_list.clear()
            pass

        record_list.append(record)
        last_pin = pin
    # 完成的单人考勤记录，排序后加入总表
    #record_list.sort(key=lambda x:x[2])
    data_dic[last_pin] = record_list

    print_dict(data_dic)
    return 1
    #end read_input_xls()
def print_dict(data_dict):
    # 输出总表测试代码
    debug_print("---------------------------------------------")
    # 建立索引表以便排序
    pin_list = []
    for pin in data_dict:
        pin_list.append(pin)
    pin_list.sort()

    for pin in pin_list:
        debug_print("pin:{0}".format(pin))
        data_list = data_dict[pin]
        for item in data_list:
            #date_tuple = xlrd.xldate_as_tuple(item.check_time, 0)
            #checktime = datetime(*date_tuple)
            #timestr = "{:%Y-%m-%d %H:%M:%S}".format(checktime)
            #debug_print("{0} [{1}]".format(item, timestr))
            debug_print(item)

    debug_print("---------------------------------------------")
    #end print_data_dict()
def write_output_xls(output_file):
    '''把内部数据列表输出到指定格式表格'''
    #--- EXAMPLE ---
    #style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold
    #on',
    #    num_format_str='#,##0.00')
    #style1 = xlwt.easyxf(num_format_str='YYYY/M/D hh:mm:ss')

    #ws.write(0, 0, 1234.56, style0)
    #ws.write(1, 0, datetime.now(), style1)
    #ws.write(2, 0, 1)
    #ws.write(2, 1, 1)
    #ws.write(2, 2, xlwt.Formula("A3+B3"))

    wb = xlwt.Workbook()
    ws = wb.add_sheet('Sheet 1')

    # 写标题栏列名
    row = 0
    style_title = xlwt.easyxf("pattern: pattern solid, fore_colour aqua; font: bold on; align: horiz center")
    xlwt.Pattern.SOLID_PATTERN
    for col in range(OutputRecord.COLUMN_COUNT):
        ws.write(row, col, OutputRecord.COLUMN_NAME[col], style_title)
        pass
    row += 1

    # 输出格式
    style_weekday = xlwt.easyxf(num_format_str = "aaaa")
    style_date = xlwt.easyxf(num_format_str = "YYYY/M/D")
    style_time = xlwt.easyxf(num_format_str = "hh:mm")
    style_blank_fill_weekend = xlwt.easyxf("pattern: pattern solid, fore_colour " + XLS_BLANK_FILL_WEEKEND)
    style_blank_fill_except  = xlwt.easyxf("pattern: pattern solid, fore_colour " + XLS_BLANK_FILL_EXCEPT)

    # 建立索引表以便排序
    pin_list = []
    for pin in data_dic:
        pin_list.append(pin)
    pin_list.sort()

    # 写数据
    for pin in pin_list:
        record_list = process_record_list(data_dic[pin])
        for record in record_list:
            ws.write(row, OutputRecord.PIN, record.pin)
            ws.write(row, OutputRecord.WEEKDAY, record.weekday, style_weekday)
            ws.write(row, OutputRecord.NAME, record.name)
            ws.write(row, OutputRecord.CHECK_DATE, record.check_date, style_date)

            if record.check_in != None:
                ws.write(row, OutputRecord.CHECK_IN, record.check_in, style_time)
            else:
                # 时间空格按是否周末填充颜色
                style_fill = style_blank_fill_except
                if record.check_date.weekday() > 4:
                    style_fill = style_blank_fill_weekend
                ws.write(row, OutputRecord.CHECK_IN, "", style_fill)

            if record.check_out != None:
                ws.write(row, OutputRecord.CHECK_OUT, record.check_out, style_time)
            else:
                # 时间空格按是否周末填充颜色
                style_fill = style_blank_fill_except
                if record.check_date.weekday() > 4:
                    style_fill = style_blank_fill_weekend
                ws.write(row, OutputRecord.CHECK_OUT, "", style_fill)

            # 判断漏打下班卡
            if record.check_in != None and record.check_out == None:
                record.except_code = 2
            if record.except_code != 0:
                ws.write(row, OutputRecord.EXCEPT_CODE,
                        OutputRecord.EXCEPT_CODE_TEXT[record.except_code],
                        xlwt.easyxf("font: colour red"))

            row += 1

    wb.save(output_file)
    return 1
    #end write_output_xls()
def get_input_checktime(in_time):
    '''不知为何时间有时读出来是Excel日期的浮点数型，有时又是字符串，干脆两个都处理为datetime'''

    out_time = None
    if type(in_time) == float:
        #debug_print("Excel float datetime")
        time_tuple = xlrd.xldate_as_tuple(in_time, 0)
        out_time = datetime(*time_tuple)
    elif type(in_time) == str:
        #debug_print("string datetime")
        out_time = datetime.strptime(in_time, "%Y/%m/%d %H:%M:%S")
    else:
        out_time = in_time

    #debug_print(out_time)
    return out_time
    #end get_input_checktime()

def process_record_list(input_record_list):
    '''
    用途：处理数据列表，参数传入输入格式列表，返回输出格式列表
    要点：
    1.按照指定列格式输出；
    2.两条上下班签到/签退记录合并为一条考勤记录；
    3.填充没有考勤记录的空白天数；
    4.注意个别的通宵加班情况，签退时间可能在第二天凌晨。
    5.如果前一天有上班，第二天6点前出现打卡记录，视为通宵加班。
    6.更为特殊的情况，比如加班跨自然月、通宵到第二天上班时间打卡等，特殊标记后人工处理。
    '''
    output_list = []
    if len(input_record_list) == 0:
        return output_list

    # 输入列表按时间排列

    # 先填充当月全部日期无签到退时间的记录
    in_data = input_record_list[0]
    checktime = in_data.check_time
    month_range = calendar.monthrange(checktime.year, checktime.month)   #返回指定月份(第一天是星期几, 该月多少天)
    for i in range(month_range[1]):
        day = datetime(checktime.year, checktime.month, i+1)
        fill_data = OutputRecord(in_data.pin, day, in_data.name, day)
        output_list.append(fill_data)

    # 遍历输入资料列表，提取签到和签退时间
    try:
        for i in range(len(input_record_list)):
            in_data = input_record_list[i]
            checktime = in_data.check_time

            # EXCEPT_CODE_TEXT = ["","上班未签到","下班未签退","通宵加班","上月末通宵加班"]
            # 以日期为索引，直接填充输出列表中对应位置。日期从1开始，索引从0
            # 先填充为空的签到时间
            if output_list[checktime.day-1].check_in == None:
                # 特殊处理：漏打上班卡，直接写入下班时间
                if checktime.time() >= WORK_OFF_TIME:
                    output_list[checktime.day-1].check_out = checktime
                    output_list[checktime.day-1].except_code = 1
                # 特殊处理：通宵加班，视为前一天下班签退
                elif checktime.time() < WHOLE_NIGHT_WORK_TIME:
                    # 更加特殊的情况：当月第一天第一条记录就是通宵加班，因没法写入上月记录，放弃自动处理，仅记录异常代码
                    if checktime.day == 1:
                        output_list[checktime.day-1].except_code = 4
                    else:
                        output_list[checktime.day-2].check_out = checktime
                        output_list[checktime.day-2].except_code = 3
                # 正常情况：当天第一条记录写入签到时间
                else:
                    output_list[checktime.day-1].check_in = checktime
            # 已有签到时间，最有可能的就是签退记录。填充签退时间
            elif output_list[checktime.day-1].check_out == None:
                output_list[checktime.day-1].check_out = checktime
            # 签到和签退时间都有，那么可能是有多条签退记录。保留最晚
            else:
                output_list[checktime.day-1].check_out = checktime
    except Exception as e:
        print(e)

    return output_list
    #end process_record_list()

def process_cmdline_args():
    '''获取命令行参数，输入的源考勤记录文件路径'''
    parser = argparse.ArgumentParser(description="把考勤机导出的考勤记录文件，转换为公司考勤表格式")
    parser.add_argument("source_file", help="源考勤记录文件路径")
    args = parser.parse_args()
    return str(args.source_file)
    #end process_cmdline_args()

#------------------
#silent_mode = True
source_path = process_cmdline_args()
if len(source_path.strip()) < 2:
    print("[ERR]Invaild source file path!")

#result = read_input_xls(r'D:\Projects\Python\AttendFormator\InData.xls')
#result = read_input_xls(r'D:\Projects\Python\AttendFormator\201707考勤记录.xls')
result = read_input_xls(source_path)

# 输出文件名为 原文件名_OutData.xls
suffix_pos = source_path.rfind('.')
if(suffix_pos > 0):
    output_file = source_path[0:suffix_pos] + "_OutData" + source_path[suffix_pos:]
    result = write_output_xls(output_file)
    print("输出考勤记录到：" + output_file)
