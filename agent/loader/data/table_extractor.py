import re
import os
import pandas as pd
import json
from tqdm import tqdm


class TableProcessor:
    def __init__(self, txt_path):
        self.txt_path = txt_path
        self.all_data = []
        self.all_table = []
        self.all_title = []

    def read_file(self):
        # 读取txt文件中的内容并将其存储在self.all_data列表中
        with open(self.txt_path, "r") as file:
            for line in file:
                data = eval(line)
                # 忽略页眉和页脚，并且内容不为空的部分
                if data["type"] not in ['页眉', '页脚'] and data["inside"] != '':
                    self.all_data.append(data)

    def process_text_data(self):
        """
        处理文本数据，提取一级标题、二级标题和表格数据
        """
        for i in range(len(self.all_data)):
            data = self.all_data[i]
            inside_content = data.get("inside")
            content_type = data.get("type")
            if content_type == "text":
                # 使用正则表达式匹配一级标题、二级标题和一级标题数字的格式
                first_level_match = re.match(
                    r"§(\d+)+([\u4e00-\u9fa5]+)", inside_content.strip()
                )
                second_level_match = re.match(
                    r"(\d+\.\d+)([\u4e00-\u9fa5]+)", inside_content.strip()
                )
                first_num_match = re.match(r"^§(\d+)$", inside_content.strip())
                # 遍历标题列表，获取所有一级标题
                title_name = [
                    dictionary["first_title"]
                    for dictionary in self.all_title
                    if "first_title" in dictionary
                ]
                if first_level_match:
                    first_title_text = first_level_match.group(2)
                    first_title_num = first_level_match.group(1)
                    first_title = first_title_num + first_title_text
                    # 如果一级标题不在已有标题列表中，则添加到标题列表中
                    if first_title not in title_name:
                        if (
                            int(first_title_num) == 1
                            or int(first_title_num) - int(self.all_title[-1]["id"]) == 1
                        ):
                            current_entry = {
                                "id": first_title_num,
                                "first_title": first_title,
                                "second_title": [],
                                "table": [],
                            }
                            self.all_title.append(current_entry)

                elif second_level_match:
                    second_title_name = second_level_match.group(0)
                    second_title = second_level_match.group(1)
                    first_title = second_title.split(".")[0]
                    if int(first_title) - 1 >= len(self.all_title):
                        continue
                    else:
                        titles = [
                            sub_item['title']
                            for sub_item in self.all_title[int(first_title) - 1][
                                'second_title'
                            ]
                        ]
                        if second_title_name not in titles:
                            self.all_title[int(first_title) - 1]["second_title"].append(
                                {"title": second_title_name, "table": []}
                            )
                elif first_num_match:
                    first_num = first_num_match.group(1)
                    first_text = self.all_data[i + 1].get("inside")
                    first_title = first_num_match.group(1) + first_text
                    # 如果标题不含有"..."，且不在已有标题列表中，则添加到标题列表中
                    if "..." not in first_text and first_title not in title_name:
                        if (
                            int(first_num) == 1
                            or int(first_num) - int(self.all_title[-1]["id"]) == 1
                        ):
                            current_entry = {
                                "id": first_num,
                                "first_title": first_title,
                                "second_title": [],
                                "table": [],
                            }
                            self.all_title.append(current_entry)

    def process_excel_data(self):
        """
        处理表格数据，提取表格内容和标题
        """
        temp_table = []
        temp_title = None

        for i in range(len(self.all_data)):
            data = self.all_data[i]
            inside_content = data.get("inside")
            content_type = data.get("type")
            if content_type == "excel":
                temp_table.append(inside_content)
                if temp_title is None:
                    for j in range(i - 1, -1, -1):
                        if self.all_data[j]["type"] == "excel":
                            break
                        if self.all_data[j]["type"] == "text":
                            content = self.all_data[j]["inside"]
                            if re.match(r'^\d+\.\d+', content) or content.startswith(
                                "§"
                            ):
                                temp_title = content.strip()
                                break
            elif content_type == "text" and temp_title is not None:
                self.all_table.append({"title": temp_title, "table": temp_table})
                temp_title = None
                temp_table = []

    def process_tables(self):
        """
        处理表格数据，将其与标题对应
        """
        for table in self.all_table:
            title = table["title"]
            table_content = table["table"]
            first_match = re.match(r'§(\d+)(.+)', title)
            second_match = re.match(r'(\d+)\.(\d+)(.+)', title)
            try:
                if first_match:
                    first_title = first_match.group(1)
                    text_part = first_match.group(2)
                    table_pair = {
                        "table_name": text_part,
                        "table_content": table_content,
                    }
                    # 遍历self.all_title列表
                    for item in self.all_title:
                        # 如果找到'id'为10的字典，则向其'second_title'列表中添加新的字典
                        if item['id'] == first_title:
                            item["table"].append(table_pair)
                            break

                elif second_match:
                    index = 0
                    for char in title:
                        if not char.isdigit() and char != '.':
                            break
                        index += 1
                    table_name = title[index:]
                    first_title, second_title = (
                        second_match.group(1),
                        int(second_match.group(2)) - 1,
                    )
                    table_pair = {
                        "table_name": table_name,
                        "table_content": table_content,
                    }
                    for item in self.all_title:
                        if item['id'] == first_title:
                            item['second_title'][second_title]["table"].append(
                                table_pair
                            )
            except:
                # 错误处理：打印错误信息并终止循环
                print(
                    "Error: 文件{}中的标题{}有误，截断处理".format(self.txt_path, title)
                )
                break

    def create_excel_files(self, output_folder):
        for item in self.all_title:
            first_title = item['first_title']
            second_title = item['second_title']
            folder_path = os.path.join(output_folder, first_title)
            os.makedirs(folder_path, exist_ok=True)
            first_title_table_data = item['table']
            try:
                if first_title_table_data != []:
                    for table_item in first_title_table_data:
                        table_name = table_item['table_name']
                        excel_name = f"{table_name}.xlsx"
                        excel_path = os.path.join(folder_path, excel_name)
                        table_content = table_item['table_content']
                        table_content = [eval(item) for item in table_content]
                        # 将第一个子列表作为列名，其余子列表作为数据行
                        max_cols = len(table_content[0])  # 第一行的列数即为最大列数
                        for row in table_content:
                            if len(row) > max_cols:  # 如果当前行的列数超过了最大列数
                                for i in range(max_cols, len(row)):
                                    row[max_cols - 1] += (
                                        ',' + row[i]
                                    )  # 合并多余的值到当前行的最大列数
                                del row[max_cols:]

                        # 创建 DataFrame
                        df = pd.DataFrame(table_content[1:], columns=table_content[0])
                        df.to_excel(excel_path, index=False)
                else:
                    for table_item in second_title:
                        second_folder_path = os.path.join(
                            folder_path, table_item["title"]
                        )
                        os.makedirs(second_folder_path, exist_ok=True)
                        for table in table_item['table']:
                            table_name = table['table_name'].replace("/", "或")
                            table_content = table['table_content']
                            table_content = [eval(item) for item in table_content]
                            excel_name = f"{table_name}.xlsx"
                            excel_path = os.path.join(second_folder_path, excel_name)
                            max_cols = len(table_content[0])  # 第一行的列数即为最大列数
                            for row in table_content:
                                if (
                                    len(row) > max_cols
                                ):  # 如果当前行的列数超过了最大列数
                                    for i in range(max_cols, len(row)):
                                        row[max_cols - 1] += (
                                            ',' + row[i]
                                        )  # 合并多余的值到当前行的最大列数
                                    del row[max_cols:]
                            # 创建 DataFrame
                            df = pd.DataFrame(
                                table_content[1:], columns=table_content[0]
                            )
                            df.to_excel(excel_path, index=False)
            except:
                # 匹配含有年月日的列名
                print(
                    "Error: 文件<{}>中的表格<{}>有误，拆分处理".format(
                        self.txt_path, table_name
                    )
                )
                pattern = r'\d{4}年\d{1,2}月\d{1,2}日'
                # 遍历第一行的列名
                for i, column_name in enumerate(table_content[0]):
                    # 如果列名匹配到含有年月日的格式
                    if re.search(pattern, column_name):
                        # 合并当前列名和左右相邻的列名
                        table_content[0][i - 1] = ' '.join(
                            [column_name, table_content[0][i - 1]]
                        )
                        table_content[0][i + 1] = ' '.join(
                            [column_name, table_content[0][i + 1]]
                        )
                        # 删除当前列及左右相邻的列
                        del table_content[0][i]
                print(table_content)
                df = pd.DataFrame(table_content[1:], columns=table_content[0])
                df.to_excel(excel_path, index=False)

        all_title_path = os.path.join(output_folder, "all_data.json")
        all_table_path = os.path.join(output_folder, "all_table.json")
        with open(all_table_path, "w", encoding="utf-8") as f:
            json.dump(self.all_table, f, ensure_ascii=False, indent=4)
        with open(all_title_path, "w", encoding="utf-8") as f:
            json.dump(self.all_title, f, ensure_ascii=False, indent=4)


# 设置txt文件夹路径
txt_folder = "./data/txt"
output_path = "./data/excel"

# 获取txt文件夹中的所有txt文件名
txt_files = [file for file in os.listdir(txt_folder) if file.endswith('.txt')]

# 对每个txt文件进行批量处理
for txt_file in tqdm(txt_files, desc="Processing txts"):
    # 构建txt文件路径
    txt_path = os.path.join(txt_folder, txt_file)

    # 使用txt文件名创建文件夹
    folder_name = txt_file.split('.')[0]
    output_folder = os.path.join(output_path, folder_name)
    os.makedirs(output_folder, exist_ok=True)
    print(txt_path)
    # 创建TableExtractor实例并处理txt文件
    processor = TableProcessor(txt_path)
    processor.read_file()
    processor.process_text_data()
    processor.process_excel_data()
    processor.process_tables()
    processor.create_excel_files(output_folder)

print("批量处理完成！")
