import os, sys

root_path = os.path.dirname(os.path.dirname(__file__))
sys.path.append(root_path)

"--------------------------------------------------xmind to excel------------------------------------------------------"
"默认配置文件"
default_path=os.path.dirname(__file__)+r"\default.ini"

source_folder_path=os.path.join(root_path,"source_folder")

"------------xmind预制文件------------------"
xmind_name = "demo.xmind"
xmind_path =os.path.join(source_folder_path,xmind_name)
"--------------测试用例模板文件---------------"
case_file = "测试用例录入模板.xlsx"
source_template_file =os.path.join(source_folder_path,case_file)
"-----------------功能点文件模板--------------"
point_file = "功能点录入模板.xlsx"
source_point_file =os.path.join(source_folder_path,point_file)

"--------------------------------------------------excel to xmind------------------------------------------------------"
"------------excel预制文件------------------"
excel_to_xmind="excel_to_xmind-demo.xlsx"       #excel_to_xmind-demo
excel_to_xmind_path=os.path.join(source_folder_path,excel_to_xmind)


"---------------------------------------------------通用配置(默认源文件名称)------------------------------------------------"
"目标路径"
temp_folder=os.path.join(root_path,"temp_folder")
"目标测试用例文件"
target_case_file =os.path.join(temp_folder,case_file)
"目标功能点文件"
target_point_file =os.path.join(temp_folder,point_file)
"目标xmind文件"
target_xmind_file=os.path.join(temp_folder,f"{excel_to_xmind[:-5]}.xmind")

if __name__ == "__main__":
    print("default_path:",default_path)
    print("temp_folder:",temp_folder)
    print("source_point_file:", source_point_file, "\ntarget_case_file:", target_case_file)