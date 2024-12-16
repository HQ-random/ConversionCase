from loguru import logger
import os

temp_path=os.path.dirname(os.path.dirname(__file__))
log=rf"{temp_path}\my_log.log"

# 配置日志格式
logger.add(log,
          rotation="10 MB",  # 滚动日志大小
          retention="1 week",  # 保留日志时间
          compression="zip",  # 压缩日志
          enqueue=True,  # 使用队列，避免阻塞
          format="{time} {level} {message}")  # 日志格式

# 配置控制台输出日志格式
# logger.add(sink=lambda msg: print(f"[{msg}]"), format="{level} {message}")

if __name__=="__main__":
    print("temp_path:", temp_path,"\tlog:",log)