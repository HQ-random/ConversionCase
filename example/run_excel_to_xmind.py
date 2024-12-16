from loguru import logger
from config import setting
from common import read_excel
from common import write_xmind


def run():
    "excel转xmind文件"

    "excel文件路径"
    excel_file_path = setting.excel_to_xmind_path
    "xmind文件路径"
    target_xmind_file = setting.target_xmind_file

    # --------------------excel数据提取，并重新调整数据结构------------------
    xdata = read_excel.Read_excel()
    data = xdata.process_data(excel_file_path)

    # --------------------对excel提取后的数据进行写入xmind中----------------
    w = write_xmind.Write_xmind()
    w.excel_to_xmind(data, target_xmind_file)

    logger.add("../my_log.log")  # 将日志输出到 my_log.log 文件


if __name__ == "__main__":
    run()
