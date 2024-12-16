import os
import time
from config import setting
import shutil
import xmind
from datetime import datetime


class Write_need_xmind:

    def __init__(self, num="3"):
        if str(num) == "3":  #excel -> 定制xmind格式
            "如果文件夹存在，则删除,清除缓存数据"
            if os.path.exists(setting.temp_folder):
                "文件夹删除"
                shutil.rmtree(setting.temp_folder)
            time.sleep(0.5)
            if not os.path.exists(setting.temp_folder):
                "创建文件夹"
                os.mkdir(setting.temp_folder)
                self.xmind_file = setting.target_xmind_file
        else:  #标准xmind -> excel -> 定制xmind格式
            num = datetime.now().strftime("%y%m%d")
            self.xmind_file = os.path.join(setting.temp_folder, f"need_xmind_{num}.xmind")

    def create_topic(self, xmind_workbook, title, num):
        "创建主题并设置主题名称"
        topic = xmind_workbook.createTopic()
        topic.setTitle(title)
        topic.addMarker(f"priority-{num}")
        return topic

    def excel_to_xmind(self, data):

        "加载一个空白的xmind文件作为模板"
        xmind_workbook = xmind.load("blank.xmind")
        if xmind_workbook is None:
            print("无法加载空白模板文件，请确保‘blank.xmind’文件存在！")
            return None

        "获取第一个工作表"
        sheet = xmind_workbook.getSheets()[0]
        "获取根主题的标题"
        root_topic = sheet.getRootTopic()
        "设置根主题的标题"
        root_topic.setTitle(data["title"])

        for point1 in data["topics"]:  #优先级1，自定义选项"一级菜单”
            point1_topic = self.create_topic(xmind_workbook, point1["title"], 1)
            root_topic.addSubTopic(point1_topic)
            for point2 in point1["topics"]:  #优先级2，自定义项“二级菜单”
                point2_topic = self.create_topic(xmind_workbook, point2["title"], 2)
                point1_topic.addSubTopic(point2_topic)
                for point3 in point2["topics"]:  #优先级3，测试项
                    point3_topic = self.create_topic(xmind_workbook, point3["title"], 3)
                    point2_topic.addSubTopic(point3_topic)
                    for point4 in point3["topics"]:  #优先级4，测试子项
                        if point4["title"] != "None":
                            point4_topic = self.create_topic(xmind_workbook, point4["title"], 4)
                            point3_topic.addSubTopic(point4_topic)

                            point5 = point4["topics"][0]["title"]  #优先级5，测试步骤
                            point5_topic = self.create_topic(xmind_workbook, point5, 5)
                            point4_topic.addSubTopic(point5_topic)

                        else:
                            point5 = point4["topics"][0]["title"]
                            point5_topic = self.create_topic(xmind_workbook, point5, 5)
                            point3_topic.addSubTopic(point5_topic)

                        point6 = point4["topics"][0]["topics"][0]["title"]  # 优先级6，预期结果
                        point6_topic = self.create_topic(xmind_workbook, point6, 6)
                        point5_topic.addSubTopic(point6_topic)
        xmind.save(xmind_workbook, self.xmind_file)


if __name__ == "__main__":
    w = Write_need_xmind()
