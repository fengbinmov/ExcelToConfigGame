@echo off

rem searchOption 0	//TopDirectoryOnly
rem searchOption 1	//AllDirectories
rem sourceDirectory	//源文件(Excel.xlsx) 目录;为空则为当前bat目录
rem targetDirectory	//输出文件的目录;为空则为当前bat目录 
rem jsonformat 0	//json 输出不格式化
rem program MultFileLanguage|Json	//选择不同的程序功能
rem autoExit 1		//自动关闭
rem jsonGroup		//excel 中存在#group 行则进行过滤，只转换与jsonGroup名字相同的列

title excel to game config

set sourceDirectory=%CD%/多语言/multilingual.xlsx
set targetDirectory=%CD%/../Language

ExcelToConfigGame sourceDirectory %sourceDirectory% targetDirectory %targetDirectory% searchOption 1 program MultFileLanguage

@echo on