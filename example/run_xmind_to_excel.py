import time
from common import read_xmind
from common import write_excel
from config import setting
from config import logger_config


def run():
    "xmind转excel运行文件"

    "xmind文件路径"
    xmind_file_path = setting.xmind_path

    # --------------------xmind数据采集、提取------------------
    xdata = read_xmind.Read_xmind()
    data_list = xdata.read_xmind_data(xmind_file_path)
    # print("data_list:",data_list)

    case_info = xdata.xmind_to_caselist(data_list)
    # print("case_info:", case_info)

    total_cases = xdata.update_case(case_info)
    # print("total_cases=", total_cases)
    logger_config.logger.info(f"total_cases={total_cases}")
    time.sleep(2)

    # ---------------对提取后的数据进行写入excel------------------
    w = write_excel.Write_excel()
    "写入测试用例数据"
    w.write_excel_case(total_cases)
    "写入功能点数据"
    w.write_excel_point(total_cases)

    logger_config.logger.info("运行完成")


if __name__ == "__main__":
    run()
