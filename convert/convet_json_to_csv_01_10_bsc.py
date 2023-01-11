import datetime
import json
import csv

import pytz

"""
将json文件数据读取后保存为csv
"""


def convert(source_file_name, to_file_name):
    """
    将json文件数据读取后保存为csv
    :param source_file_name: json文件名称
    :param csv_header: csv表头
    :return:
    """
    file = open(source_file_name, 'rb')
    data = json.loads(file.read())
    with open(to_file_name, 'w', encoding='utf-8', newline='') as fp:
        writer = csv.writer(fp)
        property_name_list = list()

        for item in data.get("data").get("get_execution").get("execution_succeeded").get("data"):
            if len(property_name_list) <= 0:
                for key_name in item.keys():
                    property_name_list.append(key_name)
                # 设置第一行标题头
                writer.writerow(property_name_list)
            row = [[]]
            for key_name in property_name_list:
                value = str(item.get(key_name))
                if "_col0" in key_name or "_col1" in key_name:
                    value = "0x" + value
                if "block_time" in key_name:
                    value = value.replace(".000 UTC", "")
                    dt = datetime.datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
                    eta = (dt + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
                    value = str(eta)
                row[0].append(value)
            # 将数据写入
            writer.writerows(row)


if __name__ == '__main__':
    source_file_name = "../data/source/dune/01-11-bsc.json"
    to_file_name = "../data/output/dune/01-11-output-bsc.csv"
    convert(source_file_name, to_file_name)
