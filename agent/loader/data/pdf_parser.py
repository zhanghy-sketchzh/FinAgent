import pdfplumber
import re
import os
from collections import defaultdict
import json
from multiprocessing import Pool


class PDFProcessor:
    def __init__(self, filepath):
        self.filepath = filepath
        self.pdf = pdfplumber.open(filepath)
        self.all_text = defaultdict(dict)
        self.allrow = 0
        self.last_num = 0

    def check_lines(self, page, top, buttom):
        lines = page.extract_words()[::]
        text = ""
        last_top = 0
        last_check = 0
        for l in range(len(lines)):
            each_line = lines[l]
            check_re = "(?:。|；|单位：人民币元|金额单位：人民币元|单位：万元|币种：人民币|\d|报告(?:全文)?(?:（修订版）|（修订稿）|（更正后）)?)$"
            if top == "" and buttom == "":
                if abs(last_top - each_line["top"]) <= 2:
                    text = text + each_line["text"]
                elif (
                    last_check > 0
                    and (page.height * 0.9 - each_line["top"]) > 0
                    and not re.search(check_re, text)
                ):

                    text = text + each_line["text"]
                else:
                    text = text + "\n" + each_line["text"]
            elif top == "":
                if each_line["top"] > buttom:
                    if abs(last_top - each_line["top"]) <= 2:
                        text = text + each_line["text"]
                    elif (
                        last_check > 0
                        and (page.height * 0.85 - each_line["top"]) > 0
                        and not re.search(check_re, text)
                    ):
                        text = text + each_line["text"]
                    else:
                        text = text + "\n" + each_line["text"]
            else:
                if each_line["top"] < top and each_line["top"] > buttom:
                    if abs(last_top - each_line["top"]) <= 2:
                        text = text + each_line["text"]
                    elif (
                        last_check > 0
                        and (page.height * 0.85 - each_line["top"]) > 0
                        and not re.search(check_re, text)
                    ):
                        text = text + each_line["text"]
                    else:
                        text = text + "\n" + each_line["text"]
            last_top = each_line["top"]
            last_check = each_line["x1"] - page.width * 0.85

        return text

    def drop_empty_cols(self, data):
        # 删除所有列为空数据的列
        transposed_data = list(map(list, zip(*data)))
        filtered_data = [
            col for col in transposed_data if not all(cell == "" for cell in col)
        ]
        result = list(map(list, zip(*filtered_data)))
        return result

    def extract_text_and_tables(self, page):
        buttom = 0
        tables = page.find_tables()
        if len(tables) >= 1:
            count = len(tables)
            for table in tables:
                # 处理越界的表格
                if table.bbox[3] < buttom:
                    pass
                else:
                    count -= 1
                    # 处理表格上方的文本
                    top = table.bbox[1]
                    text = self.check_lines(page, top, buttom)
                    text_list = text.split("\n")
                    for _t in range(len(text_list)):
                        self.all_text[self.allrow] = {
                            "page": page.page_number,
                            "allrow": self.allrow,
                            "type": "text",
                            "inside": text_list[_t],
                        }
                        self.allrow += 1

                    # 处理表格
                    buttom = table.bbox[3]
                    new_table = table.extract()
                    r_count = 0
                    for r in range(len(new_table)):
                        row = new_table[r]
                        if row[0] is None:
                            r_count += 1
                            for c in range(len(row)):
                                if row[c] is not None and row[c] not in ["", " "]:
                                    if new_table[r - r_count][c] is None:
                                        new_table[r - r_count][c] = row[c]
                                    else:
                                        new_table[r - r_count][c] += row[c]
                                    new_table[r][c] = None
                        else:
                            r_count = 0

                    end_table = []
                    end_table = []
                    for row in new_table:
                        if row[0] != None:
                            cell_list = []
                            cell_check = False
                            for cell in row:
                                if cell != None:
                                    cell = cell.replace("\n", "")
                                else:
                                    cell = ""
                                if cell != "":
                                    cell_check = True
                                cell_list.append(cell)
                            if cell_check == True:
                                end_table.append(cell_list)

                    end_table = self.drop_empty_cols(end_table)

                    # # 处理列名为空的情况
                    for i in range(len(end_table[0])):
                        if end_table[0][i] == "":
                            if 0 < i < len(end_table[0]) - 1:
                                # 左列名
                                left_column = end_table[0][i - 1]
                                # 右列名
                                right_column = end_table[0][i + 1]
                                # 修改当前列名为左列名和右列名的拼接
                                end_table[0][i] = left_column + right_column
                            else:
                                # 当前列为空且位于首列，赋值为右列名
                                # 当前列为空且位于末尾列，赋值为左列名
                                end_table[0][i] = (
                                    end_table[0][i - 1]
                                    if i == len(end_table[0]) - 1
                                    else end_table[0][i + 1]
                                )

                    # 处理列值为空的情况, 取左边的列
                    for i in range(1, len(end_table)):
                        for j in range(len(end_table[i])):
                            if end_table[i][j] == "":
                                end_table[i][j] = end_table[i][j - 1]

                    for row in end_table:
                        self.all_text[self.allrow] = {
                            "page": page.page_number,
                            "allrow": self.allrow,
                            "type": "excel",
                            "inside": str(row),
                        }
                        # self.all_text[self.allrow] = {'page': page.page_number, 'allrow': self.allrow, 'type': 'excel',
                        #                               'inside': ' '.join(row)}
                        self.allrow += 1

                    if count == 0:
                        text = self.check_lines(page, "", buttom)
                        text_list = text.split("\n")
                        for _t in range(len(text_list)):
                            self.all_text[self.allrow] = {
                                "page": page.page_number,
                                "allrow": self.allrow,
                                "type": "text",
                                "inside": text_list[_t],
                            }
                            self.allrow += 1

        else:
            text = self.check_lines(page, "", "")
            text_list = text.split("\n")
            for _t in range(len(text_list)):
                self.all_text[self.allrow] = {
                    "page": page.page_number,
                    "allrow": self.allrow,
                    "type": "text",
                    "inside": text_list[_t],
                }
                self.allrow += 1

        first_re = "[^计](?:报告(?:全文)?(?:（修订版）|（修订稿）|（更正后）)?)$"
        end_re = "^(?:\d|\\|\/|第|共|页|-|_| ){1,}"
        if self.last_num == 0:
            try:
                first_text = str(self.all_text[1]["inside"])
                end_text = str(self.all_text[len(self.all_text) - 1]["inside"])
                if re.search(first_re, first_text) and not "[" in end_text:
                    self.all_text[1]["type"] = "页眉"
                    if re.search(end_re, end_text) and not "[" in end_text:
                        self.all_text[len(self.all_text) - 1]["type"] = "页脚"
            except:
                print(page.page_number)
        else:
            try:
                first_text = str(self.all_text[self.last_num + 2]["inside"])
                end_text = str(self.all_text[len(self.all_text) - 1]["inside"])
                if re.search(first_re, first_text) and "[" not in end_text:
                    self.all_text[self.last_num + 2]["type"] = "页眉"
                if re.search(end_re, end_text) and "[" not in end_text:
                    self.all_text[len(self.all_text) - 1]["type"] = "页脚"
            except:
                print(page.page_number)

        self.last_num = len(self.all_text) - 1

    def process_pdf(self):
        for i in range(len(self.pdf.pages)):
            self.extract_text_and_tables(self.pdf.pages[i])

    def save_all_text(self, path):
        for key in self.all_text.keys():
            with open(path, "a+", encoding="utf-8") as file:
                file.write(json.dumps(self.all_text[key], ensure_ascii=False) + "\n")


def process_pdf(file_path, output_folder):
    """
    处理单个PDF文件，并将结果保存为文本文件
    """
    print("开始处理：", file_path)
    processor = PDFProcessor(file_path)
    processor.process_pdf()
    save_path = os.path.join(
        output_folder, os.path.basename(file_path).replace(".pdf", ".txt")
    )
    processor.save_all_text(save_path)
    print("完成处理：", save_path)


def process_files_in_folder(folder_path, output_folder):
    """
    处理指定文件夹中的所有PDF文件
    """
    os.makedirs(output_folder, exist_ok=True)
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".pdf"):
                file_path = os.path.join(root, file)
                process_pdf(file_path, output_folder)


if __name__ == "__main__":
    folder_path = "./data/pdf"
    output_folder = "./data/txt"
    process_files_in_folder(folder_path, output_folder)

