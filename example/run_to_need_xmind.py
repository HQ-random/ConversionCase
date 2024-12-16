import time

from config import setting
from common import read_need_excel
from common import write_need_xmind
from example import run_xmind_to_excel


def run(num):
    "excel转定制xmind文件"

    if str(num) == "3":
        "excel -> 定制xmind格式"

        # --------------------excel数据提取，并重新调整数据结构------------------
        "excel文件路径"
        excel_file_path = setting.excel_to_xmind_path
        xdata1 = read_need_excel.Read_need_excel(num)
        data = xdata1.process_data(excel_file_path)

        # --------------------对excel提取后的数据进行写入xmind中----------------
        w = write_need_xmind.Write_need_xmind(num)
        w.excel_to_xmind(data)

    else:
        "标准xmind -> excel -> 定制xmind格式"
        # --------------------标准xmind文件生成对应的excel文件------------------
        run_xmind_to_excel.run()
        time.sleep(1)
        # --------------------excel数据提取，并重新调整数据结构------------------
        xdata2 = read_need_excel.Read_need_excel(num)
        data = xdata2.process_data()

        # --------------------对excel提取后的数据进行写入xmind中----------------
        w = write_need_xmind.Write_need_xmind(num)
        w.excel_to_xmind(data)


if __name__ == "__main__":
    run(3)
