@echo off

rem searchOption 0|1		//TopDirectoryOnly(���޶���Ŀ¼) | AllDirectories(���е�Ŀ¼)
rem sourceDirectory			//Դ�ļ�(Excel.xlsx) Ŀ¼;Ϊ����Ϊ��ǰbatĿ¼
rem targetDirectory			//����ļ���Ŀ¼;Ϊ����Ϊ��ǰbatĿ¼ 
rem jsonformat -1|0|1		//json�����ʽ�� -1 ������ | 0 �����Ű� | 1 ����Ű�
rem program xxx				//ѡ��ͬ�ĳ����ܣ�MultFileLanguage --���������; Json --jsonת��;UniqueCharacter --Ψһ�ַ���ȡ(�������ļ��е��ַ���ȡ��һ���ļ���)��ԭ�ļ�Ŀ¼��Ϊ�����::���зָ�
rem extensions xxx			//��ȡ���ļ���չ��(Excel ֻ֧�� .xlsx; UniqueCharacter�п��Զ�����,Ĭ��Ϊ".txt,.xml,.json,.yml")
rem autoExit 0|1			//0 ����������ֶ��رգ�1 ����������Զ��ر�
rem jsonGroup xxx			//excel �д���#group ������й��ˣ�ֻת����xxx������ͬ����
rem unGroupDirectory xxx	//excel ������#group ����ָ�� Excel ����ļ���Ŀ¼;Ϊ������jsonGroup�����²����
rem sheetName xxx			//excel ֻת����Ӧ��Sheet�飬sheet ��Ϊ xxx
rem jsonType JArray|JObject	//Ĭ��ΪJArray�ὫExcel ���Ϊ JArray �ĸ�ʽ��JObject ������ֵ��ʽ�����趨һ��������Ϊ�ֵ��Key
rem mainkey xxx				//JObject ����������
rem startCell A1			//excel ��ȡʱ����ʼλ��(Ĭ��ΪA1)
rem endCell					//excel ��ȡʱ�Ľ���λ��(Ĭ��Ϊ����)

title excel to game config

set sourceDirectory=%CD%/�������/ģ��.xlsx
set targetDirectory=%CD%/../Client/ģ��_�������-Sheet3.json

ExcelToConfigGame sourceDirectory %sourceDirectory% targetDirectory %targetDirectory% searchOption 1 jsonformat 1 program Json sheetName Sheet3

@echo on