@echo off

rem searchOption 0	//TopDirectoryOnly
rem searchOption 1	//AllDirectories
rem sourceDirectory	//Դ�ļ�(Excel.xlsx) Ŀ¼;Ϊ����Ϊ��ǰbatĿ¼
rem targetDirectory	//����ļ���Ŀ¼;Ϊ����Ϊ��ǰbatĿ¼ 
rem jsonformat 0	//json �������ʽ��
rem program MultFileLanguage|Json	//ѡ��ͬ�ĳ�����
rem autoExit 1		//�Զ��ر�
rem jsonGroup		//excel �д���#group ������й��ˣ�ֻת����jsonGroup������ͬ����

title excel to game config

set sourceDirectory=%CD%/������/multilingual.xlsx
set targetDirectory=%CD%/../Language

ExcelToConfigGame sourceDirectory %sourceDirectory% targetDirectory %targetDirectory% searchOption 1 program MultFileLanguage

@echo on