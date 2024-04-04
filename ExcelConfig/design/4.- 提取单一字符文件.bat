@echo off

rem searchOption 0|1		//TopDirectoryOnly(仅限顶级目录) | AllDirectories(所有的目录)
rem sourceDirectory			//源文件(Excel.xlsx) 目录;为空则为当前bat目录
rem targetDirectory			//输出文件的目录;为空则为当前bat目录 
rem jsonformat -1|0|1		//json输出格式； -1 不处理 | 0 单行排版 | 1 锯齿排版
rem program xxx				//选择不同的程序功能；MultFileLanguage --多语言输出; Json --json转换;UniqueCharacter --提取单一字符文件
rem autoExit 0|1			//0 程序结束后手动关闭；1 程序结束后自动关闭
rem jsonGroup xxx			//excel 中存在#group 行则进行过滤，只转换与xxx名字相同的列
rem unGroupDirectory xxx	//excel 不存在#group 行则指定 Excel 输出文件的目录;为空则在jsonGroup过滤下不输出
rem jsonType JArray|JObject	//默认为JArray会将Excel 输出为 JArray 的格式，JObject 会输出字典格式必须设定一个主键作为字典的Key
rem mainkey xxx				//JObject 的主键定义
rem extensions .txt,.meta

title excel to game config

set sourceDirectory=%CD%/../Client
set sourceDirectory=%sourceDirectory%::%CD%/../Language
set sourceDirectory=%sourceDirectory%::%CD%/../Server
set targetDirectory=%CD%/../TextMesh-Generate/signalchars.txt

set exts=.txt,.meta,.cs,.json

ExcelToConfigGame sourceDirectory %sourceDirectory% targetDirectory %targetDirectory% searchOption 1 program UniqueCharacter extensions %exts%

@echo on