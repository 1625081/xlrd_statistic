# 北航课表xlrd统计系统

## 简介

这是一个用于统计学生信息的python文件，一共生成三个输出：

*　各时段上课人数统计(result.csv)

*　各时段上课的人(numresult.csv)

*　上各个课的人(json-result.txt)

## 安装及使用

安装xlrd:

1.找到python安装路径下Script文件夹
2.找到xls文件路径
3.打开Windows下的cmd命令行
4.输入:
```shell
cd python安装路径下Script文件夹
pip install xlrd
```
5.打开xlrd_statistic.py文件（此时"import xlrd"这一步就不会出错了）
6.修改path变量为xls文件路径
7.运行Python程序
