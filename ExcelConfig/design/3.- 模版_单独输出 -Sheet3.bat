@echo off

rem searchOption 0|1		//TopDirectoryOnly(仅限顶级目录) | AllDirectories(所有的目录)
rem sourceDirectory			//源文件(Excel.xlsx) 目录;为空则为当前bat目录
rem targetDirectory			//输出文件的目录;为空则为当前bat目录 
rem jsonformat -1|0|1		//json输出格式； -1 不处理 | 0 单行排版 | 1 锯齿排版
rem program xxx				//选择不同的程序功能；MultFileLanguage --多语言输出; Json --json转换;UniqueCharacter --唯一字符提取(将所有文件中的字符提取到一份文件中)，原文件目录可为多个用::进行分割
rem extensions xxx			//读取的文件扩展名(Excel 只支持 .xlsx; UniqueCharacter中可自定定义,默认为".txt,.xml,.json,.yml")
rem autoExit 0|1			//0 程序结束后手动关闭；1 程序结束后自动关闭
rem jsonGroup xxx			//excel 中存在#group 行则进行过滤，只转换与xxx名字相同的列
rem unGroupDirectory xxx	//excel 不存在#group 行则指定 Excel 输出文件的目录;为空则在jsonGroup过滤下不输出
rem sheetName xxx			//excel 只转换对应的Sheet组，sheet 名为 xxx
rem jsonType JArray|JObject	//默认为JArray会将Excel 输出为 JArray 的格式，JObject 会输出字典格式必须设定一个主键作为字典的Key
rem mainkey xxx				//JObject 的主键定义
rem startCell A1			//excel 读取时的起始位置(默认为A1)
rem endCell					//excel 读取时的结束位置(默认为无限)

title excel to game config

set sourceDirectory=%CD%/数据配表/模版.xlsx
set targetDirectory=%CD%/../Client/模版_单独输出-Sheet3.json

ExcelToConfigGame sourceDirectory %sourceDirectory% targetDirectory %targetDirectory% searchOption 1 jsonformat 1 program Json sheetName Sheet3

@echo on