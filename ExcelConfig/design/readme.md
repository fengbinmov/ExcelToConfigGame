## 该目录是策划配置表目录;

 - 该目录包含若干符合策划配表规则的excel表

   ​	配表规则详见/数据配表/模板.xlsx

   ​	指定的excel文件必须符合策划配表规则，否则转表工具无法生成json文件

 - 该目录包含一个转表工具可执行文件ExcelToConfigGame.exe 和 若干 .bat 程序

   - 使用方式为双击对应 bat文件即可

   - bat 文件内可指定excel目录与输出目录，以及更多设定参数
   - 1.转换多语言.bat
     	将指定的文件转换为多个不同语言的翻译文件
   - 2.转换数据配表.bat
     	将指定的文件转换为多个不同的json文件

   

## bat 文件设定参数

```
searchOption 0|1		//TopDirectoryOnly(仅限顶级目录) | AllDirectories(所有的目录)
sourceDirectory			//源文件(Excel.xlsx) 目录;为空则为当前bat目录
targetDirectory			//输出文件的目录;为空则为当前bat目录 
jsonformat -1|0|1		//json输出格式； -1 不处理 | 0 单行排版 | 1 锯齿排版
program xxx				//选择不同的程序功能；
						MultFileLanguage --多语言输出; 
						Json --json转换
						UniqueCharacter --唯一字符提取(将所有文件中的字符提取到一份文件中)，原文件目录可为多个用::进行分割
extensions xxx			//读取的文件扩展名(Excel 只支持 .xlsx; UniqueCharacter中可自定定义,默认为".txt,.xml,.json,.yml")
autoExit 0|1			//0 程序结束后手动关闭；1 程序结束后自动关闭
jsonGroup xxx			//excel 中存在#group 行则进行过滤，只转换与xxx名字相同的列
unGroupDirectory xxx	//excel 不存在#group 行则指定 Excel 输出文件的目录;为空则在jsonGroup过滤下不输出
sheetName xxx			//excel 只转换对应的Sheet组，sheet 名为 xxx
jsonType JArray|JObject	//默认为JArray会将Excel 输出为 JArray 的格式，JObject 会输出字典格式必须设定一个主键作为字典的Key
mainkey xxx				//JObject 的主键定义
startCell A1			//excel 读取时的起始位置(默认为A1)
endCell					//excel 读取时的结束位置(默认为无限)
```

