import os
import time
from config import setting
import shutil
import xmind


class Write_xmind:

    def __init__(self):
        "如果文件夹存在，则删除,清除缓存数据"
        if os.path.exists(setting.temp_folder):
            "文件夹删除"
            shutil.rmtree(setting.temp_folder)
        time.sleep(0.5)
        if not os.path.exists(setting.temp_folder):
            "创建文件夹"
            os.mkdir(setting.temp_folder)

    def create_topic(self, xmind_workbook, title):
        "创建主题并设置主题名称"
        topic = xmind_workbook.createTopic()
        topic.setTitle(title)
        return topic

    def excel_to_xmind(self, data, xmind_file):

        "加载一个空白的xmind文件作为模板"
        xmind_workbook = xmind.load("blank.xmind")
        if xmind_workbook is None:
            print("无法加载空白模板文件，请确保‘blank.xmind’文件存在！")
            return

        # "获取xmind模板的主工作表中的根主题"
        # root_topic = xmind_workbook.getPrimarySheet().getRootTopic()
        # print("root_topic:", root_topic)
        # title=data["title"]
        # "创建新主题并将其作为主题添加到跟主题下"
        # main_topic = self.create_topic(xmind_workbook,title)

        "获取第一个工作表"
        sheet = xmind_workbook.getSheets()[0]
        "获取根主题的标题"
        root_topic = sheet.getRootTopic()
        "设置根主题的标题"
        root_topic.setTitle(data["title"])

        for menu1 in data["topics"]:  #一级菜单数量
            menu1_topic = self.create_topic(xmind_workbook, menu1["title"])
            root_topic.addSubTopic(menu1_topic)
            for menu2 in menu1["topics"]:  #二级菜单数量
                menu2_topic = self.create_topic(xmind_workbook, menu2["title"])
                menu1_topic.addSubTopic(menu2_topic)
                # for purpose in menu2["topics"]:  #测试目的
                #     purpose_topic = self.create_topic(xmind_workbook, purpose["title"])
                #     menu2_topic.addSubTopic(purpose_topic)
                for test_items in menu2["topics"]:  #测试项
                    test_items_topic = self.create_topic(xmind_workbook, test_items["title"])
                    menu2_topic.addSubTopic(test_items_topic)
                    for test_child in test_items["topics"]:  #测试子项
                        if test_child["title"] != "测试子项：None":
                            test_child_topic = self.create_topic(xmind_workbook, test_child["title"])
                            test_items_topic.addSubTopic(test_child_topic)

                            step = test_child["topics"][0]["title"]  #测试步骤
                            step_topic = self.create_topic(xmind_workbook, step)
                            test_child_topic.addSubTopic(step_topic)

                            result = test_child["topics"][0]["topics"][0]["title"]  # 预期结果
                            result_topic = self.create_topic(xmind_workbook, result)
                            step_topic.addSubTopic(result_topic)

                        else:  #测试子项为空，则无需创建

                            step = test_child["topics"][0]["title"]  # 测试步骤
                            step_topic = self.create_topic(xmind_workbook, step)
                            test_items_topic.addSubTopic(step_topic)

                            result = test_child["topics"][0]["topics"][0]["title"]  # 预期结果
                            result_topic = self.create_topic(xmind_workbook, result)
                            step_topic.addSubTopic(result_topic)

                        priority = test_child["topics"][0]["topics"][0]["topics"][0]["priority"]  # 重要程度
                        if priority[5:] != "P2":
                            priority_topics = self.create_topic(xmind_workbook, priority)
                            result_topic.addSubTopic(priority_topics)

                        condition = test_child["topics"][0]["topics"][0]["topics"][1]["condition"]  # 预置条件
                        if condition[5:] != "None":
                            condition_topic = self.create_topic(xmind_workbook, condition)
                            result_topic.addSubTopic(condition_topic)

                        scene = test_child["topics"][0]["topics"][0]["topics"][2]["scene"]  # 覆盖场景
                        if scene[5:] != "None":
                            scene_topic = self.create_topic(xmind_workbook, scene)
                            result_topic.addSubTopic(scene_topic)

                        direction = test_child["topics"][0]["topics"][0]["topics"][3]["direction"]  # 正反例
                        if direction[4:] != "正例":
                            direction_topic = self.create_topic(xmind_workbook, direction)
                            result_topic.addSubTopic(direction_topic)

                        case_type = test_child["topics"][0]["topics"][0]["topics"][4]["case_type"]  # 用例类型
                        if case_type[5:] != "功能类":
                            case_type_topic = self.create_topic(xmind_workbook, case_type)
                            result_topic.addSubTopic(case_type_topic)

                        remark = test_child["topics"][0]["topics"][0]["topics"][5]["remark"]  # 备注
                        if remark[3:] != "None":
                            # print("remark=",remark,"\tremark[3:]=",remark[3:],"\ttype=",type(remark),"\tlen(remark[3:])=",len(remark[3:]))
                            remark_topic = self.create_topic(xmind_workbook, remark)
                            result_topic.addSubTopic(remark_topic)
        xmind.save(xmind_workbook, xmind_file)


if __name__ == "__main__":
    w = Write_xmind()
