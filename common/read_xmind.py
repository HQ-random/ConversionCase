import re
from xmindparser import xmind_to_dict
import configparser
from config import setting
from config import logger_config


class Read_xmind:
    def read_xmind_data(self, xmind_file_path):
        """
        读取xmind文件，返回中心主题标题和其他分支list数据
        :param xmind_file: xmind文件路径
        :return:
        """
        case_data_dict = xmind_to_dict(xmind_file_path)[0]["topic"]
        title = case_data_dict["title"]
        data_list = case_data_dict["topics"]
        return data_list

    def find_first_colon(self, text):
        """查找字符串中第一个冒号的位置，并确保其位置在第 3 到第 6 位之间。
        Args:
            text: 输入字符串。
        Returns:
            第一个冒号的位置，如果找不到或位置不在范围内，则返回 -1。
        """
        match = re.search(r"[:：]", text)
        if match and 3 <= match.start() <= 6:
            # print("match.start():", match.start())
            return match.start()
        else:
            return -1

    def find_replace(self, index_num):
        "替换xmind选填项值"
        case = {"menu1": "一级菜单", "menu2": "二级菜单", "purpose": "测试目的", "test_items": "测试项",
                "test_child": "测试子项", "step": "测试步骤", "result": "预期结果", "trading_day": "交易日",
                "condition": "预置条件", "scene": "覆盖场景", "case_num": "测试用例数", "baseline": "是否基线用例",
                "create_time": "编写日期", "priority": "重要程度", "direction": "正反例", "case_type": "用例类型",
                "create_name": "编写人", "remark": "备注"}
        case1 = {}
        colon_index = self.find_first_colon(index_num)
        if colon_index != -1:
            key = index_num[:colon_index].strip()
            # print("key:", key)
            value = index_num[colon_index + 1:].strip()
            # print("value:", value)
        else:
            key = None
            value = None

        if key and key in case.values():
            key1 = list(case.keys())[list(case.values()).index(key)]
            # print("key1", key1, "\tkey:", key)
            case1[key1] = value
            # print("case1=",case1)
        else:
            print(f"未找到匹配的键: {key}")
        return case1

    def xmind_to_caselist(self, data_list):
        """
        根据传入的list数据，递归解析出与模板一致的数据，
        :param data_list: 传入解析后list
        :param cases: 存用例数据的list，默认为空
        :return: 返回以每条用例数据的list
        """
        cases = []
        # print("{:-^100}".format("开始"))
        logger_config.logger.info("{:-^100}".format("开始"))
        for menu1 in range(len(data_list)):
            "一级菜单数量"
            # print("menu1:", len(data_list))
            # print("一级菜单名称：", data_list[menu1]["title"])
            for menu2 in range(len(data_list[menu1]["topics"])):
                "二级菜单数量"
                # print("menu2=", len(data_list[menu1]["topics"]))
                # print("二级菜单名称=", data_list[menu1]["topics"][menu2]["title"])
                # for purpose in range(len(data_list[menu1]["topics"][menu2]["topics"])):
                #     "测试目的数量——11.17改，去除'测试目的'项"
                # print("purpose:", len(data_list[menu1]["topics"][menu2]["topics"]))
                # print("测试目的名称", data_list[menu1]["topics"][menu2]["topics"][purpose]["title"])
                for test_items in range(len(data_list[menu1]["topics"][menu2]["topics"])):
                    # print("test_items:", len(data_list[menu1]["topics"][menu2]["topics"][purpose]["topics"]))
                    # print("测试项的名称",data_list[menu1]["topics"][menu2]["topics"][purpose]["topics"][test_items]["title"])
                    # print("{:-^50}".format("测试项换行"))
                    "测试项数量"
                    test_items_num = len(data_list[menu1]["topics"][menu2]["topics"][test_items]["topics"])

                    "测试子项：测试项下一节点开头匹配判断"
                    title = data_list[menu1]["topics"][menu2]["topics"][test_items]["topics"][0]["title"]
                    pattern = re.compile(r"^测试子项")
                    match = re.search(pattern, title)
                    if match:
                        "根据节点名称判断是否存在测试子项"
                        for test_child_dict in data_list[menu1]["topics"][menu2]["topics"][test_items]["topics"]:
                            # 创建一个新的字典
                            new_case = {}
                            # print("test_child_dict=",test_child_dict)
                            # new_case["test_child"] = title[5:]
                            "一级菜单名称"
                            new_case["menu1"] = data_list[menu1]["title"][5:]
                            "二级菜单名称"
                            new_case["menu2"] = data_list[menu1]["topics"][menu2]["title"][5:]
                            "测试项名称"
                            new_case["test_items"] = data_list[menu1]["topics"][menu2]["topics"][test_items]["title"][
                                                     4:]
                            "测试子项"
                            new_case["test_child"] = test_child_dict["title"][5:]
                            "测试步骤名称"
                            new_case["step"] = test_child_dict["topics"][0]["title"][5:]
                            "预期结果名称"
                            new_case["result"] = test_child_dict["topics"][0]["topics"][0]["title"][5:]
                            if test_child_dict["topics"][0]["topics"][0].get("topics", 1) != 1:
                                "选填是否存在"
                                optional_num = len(test_child_dict["topics"][0]["topics"][0]["topics"])
                                for x in range(optional_num):
                                    # new_case["optional"] = data_list[menu1]["topics"][menu2]["topics"][purpose]["topics"][test_items]["topics"][0]["topics"][0]["topics"][0]["topics"]
                                    optional = test_child_dict["topics"][0]["topics"][0]["topics"][x]["title"]
                                    # index_num=self.find_first_colon(optional)
                                    # print("index_num=",index_num,"index_num:",type(index_num))
                                    optional_info = self.find_replace(optional)
                                    new_case.update(optional_info)
                            "测试目的名称"
                            new_case[
                                "purpose"] = f"验证{new_case['menu2']}_{new_case['test_items']}_{new_case['test_child']}功能正常"
                            cases.append(new_case)
                            logger_config.logger.info("{:-^50}".format("case_data"))
                            logger_config.logger.info(f"new_case:{new_case}")

                    else:  #测试子项为空
                        new_case = {}
                        "一级菜单名称"
                        new_case["menu1"] = data_list[menu1]["title"][5:]
                        "二级菜单名称"
                        new_case["menu2"] = data_list[menu1]["topics"][menu2]["title"][5:]
                        "测试项名称"
                        new_case["test_items"] = data_list[menu1]["topics"][menu2]["topics"][test_items]["title"][4:]
                        "测试步骤名称"
                        new_case["step"] = data_list[menu1]["topics"][menu2]["topics"][test_items]["topics"][0][
                                               "title"][5:]
                        "预期结果名称"
                        new_case["result"] = \
                        data_list[menu1]["topics"][menu2]["topics"][test_items]["topics"][0]["topics"][0]["title"][5:]
                        # print(new_case["result"])
                        if data_list[menu1]["topics"][menu2]["topics"][test_items]["topics"][0]["topics"][0].get(
                                "topics", 1) != 1:
                            "选填是否存在"
                            optional_num = len(
                                data_list[menu1]["topics"][menu2]["topics"][test_items]["topics"][0]["topics"][0][
                                    "topics"])
                            for x in range(optional_num):
                                optional = \
                                data_list[menu1]["topics"][menu2]["topics"][test_items]["topics"][0]["topics"][0][
                                    "topics"][x]["title"]
                                optional_info = self.find_replace(optional)
                                new_case.update(optional_info)
                        "测试目的名称"
                        new_case["purpose"] = f"验证{new_case['menu2']}_{new_case['test_items']}功能正常"
                        # print("{:-^50}".format("case_data"))
                        # print("new_case:", new_case)
                        cases.append(new_case)
                        logger_config.logger.info("{:-^50}".format("case_data"))
                        logger_config.logger.info(f"new_case:{new_case}")

        # print("{:-^100}".format("结束"))
        logger_config.logger.info("{:-^100}".format("结束"))
        return cases

    def update_case(self, cases):
        "更新case默认数据"
        "创建configparser对象"
        info = configparser.ConfigParser()
        "读取INI文件"
        info.read(setting.default_path, encoding="utf-8")
        # print("ini_data:",info)
        # 获取所有 section 的信息
        case_info = {}
        total_cases = []
        for section in info.sections():
            for key, value in info.items(section):
                case_info[key] = value
        # print("case_info:",case_info)
        for case in cases:
            new_case = {}
            new_case.update(case_info)
            new_case.update(case)
            total_cases.append(new_case)
        return total_cases


if __name__ == "__main__":
    from config import setting

    xmind_file_path = r"D:\PythonProject\ConversionCase\source_folder\xmind转excel改.xmind"

    xdata = Read_xmind()
    data_list = xdata.read_xmind_data(xmind_file_path)
    print("data_list=", data_list)

    case_info = xdata.xmind_to_caselist(data_list)
    print("case_info=", case_info)

    print('{:-^80}'.format("换行"))
    total_cases = xdata.update_case(case_info)
    print("total_cases=", total_cases)