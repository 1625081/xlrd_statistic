# 北航课表xlrd统计系统

## 简介

这是一个用于统计学生信息的python文件，一共生成三个输出：

*　各时段上课人数统计(result.csv)

*　各时段上课的人(numresult.csv)

*　上各个课的人(json-result.txt)

## 安装及使用

如果电脑中没有安装过xlrd模块的话，import xlrd时会报错。可以按照以下步骤安装xlrd:

1.访问：https://www.python.org/downloads/windows/

下载Python 3.5.3 - 2017-01-17下的：

Windows x86-64 executable installer（64位系统）

Windows x86 executable installer （32位系统）

安装时注意勾选：

add Python3.5 to PATH

点击install now

2.找到xls（课表）文件路径

3.打开Windows下的cmd命令行

4.输入:
```shell
pip install xlrd
```
5.打开xlrd_statistic.py文件（此时"import xlrd"这一步就不会出错了）

6.修改path变量为xls文件路径

7.运行Python程序
