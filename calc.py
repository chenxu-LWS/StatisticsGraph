import matplotlib.pyplot as plt
import sys
import xlrd
import pandas as pd

output_excel = {'eventID': [], 'zero lasts': [], 'two lasts': [],
                'medium lasts': [], 'zero proportion': [], 'two proportion': [], 'medium proportion': []}
event_id_List = []
zero_lasts_list = []
two_lasts_list = []
medium_lasts_list = []
zero_proportion_list = []
two_proportion_list = []
medium_proportion_list = []

def read_xlrd_sheet1(name):
    wb = xlrd.open_workbook(name)
    # 按工作簿定位工作表
    sh = wb.sheet_by_name('Sheet1')
    row_num = sh.nrows
    # 根据事件ID聚合，生成字典
    eventDict = {}
    # 遍历excel，每个eventID对应一个dict的key
    for i in range(row_num):
        if i != 0:
            eventID = sh.cell(i, 0).value
            eventDict.setdefault(eventID, []).append(sh.row_values(i))
    return eventDict

def read_xlrd_sheet2(name):
    wb = xlrd.open_workbook(name)
    # 按工作簿定位工作表
    sh = wb.sheet_by_name('Sheet2')
    row_num = sh.nrows
    # 根据事件ID聚合
    events = {}
    # 遍历excel，每个eventID对应一个dict的key
    for i in range(row_num):
        if i != 0:
            eventID = sh.cell(i, 0).value
            events[eventID] = sh.row_values(i)
    return events

def calc_graph(x_axis_data, y_axis_data, info):
    all_time = round(info[4] * 0.1, 1)
    point_num = len(x_axis_data)
    max_time = max(x_axis_data)
    min_time = max_time - all_time
    print(min_time)
    curr_time = max_time
    i = len(x_axis_data) - 1
    zero_last = 0.0
    two_last = 0.0
    medium_last = 0.0
    while curr_time >= min_time and i >= 1:
        curr_time = x_axis_data[i]
        curr_speed = y_axis_data[i]
        past_time = x_axis_data[i-1]
        past_speed = y_axis_data[i-1]
        if curr_speed == 0 and past_speed == 0:
            zero_last += (curr_time - past_time)
        elif curr_speed == 2 and past_speed == 2:
            two_last += (curr_time - past_time)
        else:
            medium_last += (curr_time - past_time)
        i -= 1
        curr_time = x_axis_data[i]
    zero_lasts_list.append(round(zero_last, 2))
    print("0 lasts: ", round(zero_last, 2))
    two_lasts_list.append(round(two_last, 2))
    print("2 lasts: ", round(two_last, 2))
    medium_lasts_list.append(round(medium_last, 2))
    print("medium lasts: ", round(medium_last, 2))
    total = zero_last + two_last + medium_last
    zero_proportion_list.append(round(zero_last / total, 2))
    print("0 proportion: ", round(zero_last / total, 2))
    two_proportion_list.append(round(two_last / total, 2))
    print("2 proportion: ", round(two_last / total, 2))
    medium_proportion_list.append(round(medium_last / total, 2))
    print("medium proportion: ", round(medium_last / total, 2))


def generate_data(dict_sheet1):
    # 根据dict生成图中的数据
    dict_x_axis_data = {}
    dict_y_axis_data = {}
    for eventid in dict_sheet1:
        x_axis_data = [0.0]
        y_axis_data = [2.0]
        data = dict_sheet1[eventid]
        temp_x = 0.0
        temp_y = 2.0
        for row in data:
            time = int(row[3])
            category = int(row[5])
            if category == 1:
                extra_time = float(time) * 0.1 - temp_y
                # 正常减
                if extra_time < 0:
                    temp_x += time * 0.1
                    temp_y -= float(time) * 0.1
                else:
                    temp_x += temp_y
                    temp_y = 0
                x_axis_data.append(float('%.1f' % temp_x))
                y_axis_data.append(round(temp_y, 1))
                if extra_time > 0:
                    temp_y = 0
                    temp_x += extra_time
                    x_axis_data.append(float('%.1f' % temp_x))
                    y_axis_data.append(round(temp_y, 1))
            elif category == 2:
                if time <= 10:
                    temp_x += time * 0.1
                    x_axis_data.append(float('%.1f' % temp_x))
                    y_axis_data.append(round(temp_y, 1))
                else:
                    # 一秒内不变
                    temp_x += 1
                    x_axis_data.append(float('%.1f' % temp_x))
                    y_axis_data.append(round(temp_y, 1))
                    # 一秒后每秒减少1
                    extra_time = time - 10
                    ext = temp_y - extra_time * 0.1
                    # 正常减
                    if ext > 0:
                        temp_x += extra_time * 0.1
                        temp_y -= extra_time * 0.1
                    else:
                        temp_x += temp_y
                        temp_y = 0
                    x_axis_data.append(float('%.1f' % temp_x))
                    y_axis_data.append(round(temp_y, 1))
                    if ext < 0:
                        temp_x -= ext
                        temp_y = 0
                        x_axis_data.append(float('%.1f' % temp_x))
                        y_axis_data.append(round(temp_y, 1))
            elif category == 3:
                if time <= 1:
                    temp_x += time * 0.1
                    x_axis_data.append(float('%.1f' % temp_x))
                    y_axis_data.append(round(temp_y, 1))
                else:
                    # 0.1秒内不变
                    temp_x += 0.1
                    x_axis_data.append(float('%.1f' % temp_x))
                    y_axis_data.append(round(temp_y, 1))
                    # 0.1秒后每秒+1
                    extra_time = time - 1
                    ext = 2 - (temp_y + extra_time * 0.1)
                    # 正常加
                    if ext > 0:
                        temp_x += extra_time * 0.1
                        temp_y += extra_time * 0.1
                    # 超出了2
                    else:
                        temp_x += 2 - temp_y
                        temp_y = 2
                    x_axis_data.append(float('%.1f' % temp_x))
                    y_axis_data.append(round(temp_y, 1))
                    if ext < 0:
                        temp_x -= ext
                        temp_y = 2
                        x_axis_data.append(float('%.1f' % temp_x))
                        y_axis_data.append(round(temp_y, 1))
        dict_x_axis_data[eventid] = x_axis_data
        dict_y_axis_data[eventid] = y_axis_data
    return dict_x_axis_data, dict_y_axis_data


if __name__ == '__main__':
    if len(sys.argv) <= 1:
        print("Usage: python3 main.py filename\n")
        exit(-1)
    print("Read file", sys.argv[1])
    dict_sheet1 = read_xlrd_sheet1(sys.argv[1])
    dict_sheet2 = read_xlrd_sheet2(sys.argv[1])
    if len(dict_sheet1) == 0:
        print("eventID not found!\n")
        exit(-1)
    dict_x_axis_data, dict_y_axis_data = generate_data(dict_sheet1)
    print("len: ", len(dict_y_axis_data))
    print(dict_sheet2)
    for eventID in dict_sheet2:
        print("****", eventID, "****")
        if dict_x_axis_data.get(eventID) is not None and dict_y_axis_data.get(eventID) is not None:
            event_id_List.append(eventID)
            calc_graph(dict_x_axis_data[eventID], dict_y_axis_data[eventID], dict_sheet2[eventID])
    output_excel['eventID'] = event_id_List
    output_excel['zero lasts'] = zero_lasts_list
    output_excel['two lasts'] = two_lasts_list
    output_excel['medium lasts'] = medium_lasts_list
    output_excel['zero proportion'] = zero_proportion_list
    output_excel['two proportion'] = two_proportion_list
    output_excel['medium proportion'] = medium_proportion_list
    output = pd.DataFrame(output_excel)
    output.to_excel('result.xlsx')
