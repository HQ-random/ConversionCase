# ConversionCase
已实现excel与xmind相互间的转换

1、config/default.ini：
    系统名称、版本号、需求摘要字段值从“引入数据”中导入；
    交易日、预置条件、组合维度/覆盖场景、测试用例数、编写日期、是否基线、正反例、用例类型字段值设计为默认值；
    交易日、预置条件、组合维度/覆盖场景默认置空；
    测试用例数：默认1；是否基线：默认Y；编写日期：默认转换excel当天；正反例：默认正例；用例类型：默认功能类；

2、环境依赖：
    loguru          0.7.2
    XMind           1.2.0
    xmindparser     1.0.9

3、执行：
    example：单例执行
    run.py：主文件，集成所有功能
    手动修改：在config/setting.py中修改文件名即可，源文件放入到source_folder目录，生成的文件在temp_folder目录；
    
4、效果图：
    xmind to excel：![image](https://github.com/user-attachments/assets/d2317283-5858-4718-9eee-48ce7d7f11f3)
    转换后的excel：![image](https://github.com/user-attachments/assets/c45b0bd1-9151-4264-a4a6-4256e1a748b4)
    功能描述：![image](https://github.com/user-attachments/assets/d0c1421d-57fe-4506-8c6e-f06057679305)
