@echo off

rem searchOption 0|1		//TopDirectoryOnly(���޶���Ŀ¼) | AllDirectories(���е�Ŀ¼)
rem sourceDirectory			//Դ�ļ�(Excel.xlsx) Ŀ¼;Ϊ����Ϊ��ǰbatĿ¼
rem targetDirectory			//����ļ���Ŀ¼;Ϊ����Ϊ��ǰbatĿ¼ 
rem jsonformat -1|0|1		//json�����ʽ�� -1 ������ | 0 �����Ű� | 1 ����Ű�
rem program xxx				//ѡ��ͬ�ĳ����ܣ�MultFileLanguage --���������; Json --jsonת��;UniqueCharacter --��ȡ��һ�ַ��ļ�
rem autoExit 0|1			//0 ����������ֶ��رգ�1 ����������Զ��ر�
rem jsonGroup xxx			//excel �д���#group ������й��ˣ�ֻת����xxx������ͬ����
rem unGroupDirectory xxx	//excel ������#group ����ָ�� Excel ����ļ���Ŀ¼;Ϊ������jsonGroup�����²����
rem jsonType JArray|JObject	//Ĭ��ΪJArray�ὫExcel ���Ϊ JArray �ĸ�ʽ��JObject ������ֵ��ʽ�����趨һ��������Ϊ�ֵ��Key
rem mainkey xxx				//JObject ����������
rem extensions .txt,.meta

title excel to game config

set sourceDirectory=%CD%/../Client
set sourceDirectory=%sourceDirectory%::%CD%/../Language
set sourceDirectory=%sourceDirectory%::%CD%/../Server
set targetDirectory=%CD%/../TextMesh-Generate/signalchars.txt

set exts=.txt,.meta,.cs,.json

ExcelToConfigGame sourceDirectory %sourceDirectory% targetDirectory %targetDirectory% searchOption 1 program UniqueCharacter extensions %exts%

@echo on