import json
import csv

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
                if "邀请人" in key_name:
                    value = value.replace("\\x", "0x")
                row[0].append(value)
            # 将数据写入
            writer.writerows(row)


if __name__ == '__main__':
    source_file_name = "../data/source/11-11-bsc_source-3.json"
    to_file_name = "../data/output/11-11-bsc_output-3.csv"
    convert(source_file_name, to_file_name)
