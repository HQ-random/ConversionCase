from example import run_xmind_to_excel
from example import run_excel_to_xmind
from example import run_to_need_xmind

print(
    """
使用说明：
    标准xmind文件转excel：请输入数字'1';
    excel文件转标准xmind文件：请输入数字'2';
    excel文件转定制xmind文件：请输入数字'3';
    标准xmind文件 -> excel文件 -> 定制xmind文件：默认功能;
    """
)


def run(num):
    if num == "1":
        run_xmind_to_excel.run()
    elif num == "2":
        run_excel_to_xmind.run()
    else:
        run_to_need_xmind.run(num)


if __name__ == "__main__":
    num = input("请输入想要实现的功能数字：")
    run(num)