# VASS_Automation
### 作者：张阳（2024.05-11实习生）
### 最初版发布时间:2024.07.23

## 该项目主要基于python实现自动化收集整理VASS报名信息，项目已上传至我的github:
* https://github.com/DrYangZ

## 环境配置：
* python_embed 3.12
* openpyxl(主要用于对.xlsx文件及其Sheet的读写操作)
* regex(正则表达式)

## 文件结构：
'''
VASS_Automation
 ├── README.md:程序使用手册
 ├── FeeInTotal_run.bat：报考金额统计模块入口
 ├── RegistrationFromAll_run.bat：考生及公司信息统计模块入口
 ├── RegistrationTrans_run.bat：考生信息转移模块入口
 ├── AutomationScripts
 ├──├── FeeInTotal.py：报考金额统计模块脚本
 ├──├── RegistrationFromAll.py：考生及公司信息统计模块脚本
 ├──├── RegistrationTrans.py：考生信息转移模块脚本
 ├──├── Fee in total_2024Q3.xlsx：测试表格_1
 ├──├── Registration form_all_2024Q3_test.xlsx：测试表格_2
 ├──├── Registration form_all_V03.xlsx：测试表格_3
'''

## 程序使用说明
* 按照“考试流程文档”的规范收集各公司的报名信息
* 运行RegistrationFromAll_run.bat，将报名信息汇总至“Registration form_all_（对应考试季度）”
* 运行FeeInTotal_run.bat，将报名金额汇总至“Fee in total_2024Q3”
* 运行RegistrationTrans_run.bat，将报名信息转移至与VGC-Academy共享的xlsx文件

## 这是我的第一个github项目，非常荣幸能为你带来帮助，谢谢！

![Thanks](AutomationScripts/Data/Thanks.jpg)