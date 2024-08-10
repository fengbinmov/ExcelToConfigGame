using System.Text;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace ExcelToConfigGame
{
    internal class Program
    {
        public static int autoExit = 0;

        public static int FileCount = 0;
        public static int FileNow = 0;

        public static Regex regex;

        static string[] TypeSplit(string txt,string start)
        {
            if (string.IsNullOrEmpty(txt)) return null;
            if (!txt.StartsWith(start)) return null;

            txt = txt.Replace(" ", string.Empty);
            txt = txt.Remove(0, start.Length);

            List<string> item = new List<string>();
            List<string> result = new List<string>();

            StringBuilder sb = new StringBuilder();
            int ceng = 0;

            for (int i = 0; i < txt.Length; i++)
            {
                switch (txt[i])
                {
                    case ',':
                        if (ceng > 1)
                        {
                            sb.Append(txt[i]);
                        }
                        else
                        {
                            item.Add(sb.ToString());
                            sb.Clear();
                        }
                        continue;
                    case '{':
                        ceng++;
                        if(ceng > 1) sb.Append(txt[i]);
                        continue;
                    case '}':
                        if (ceng > 1) sb.Append(txt[i]);
                        ceng--;
                        continue;
                    default:
                        sb.Append(txt[i]);
                        break;
                }
            }

            if (sb.Length > 0)
            {
                item.Add(sb.ToString());
                sb.Clear();
            }

            foreach (var ite in item)
            {
                result.AddRange(ite.Split(':', 2));
            }

            return result.ToArray();
        }

        static void Main(string[] args)
        {
            regex = new Regex("(?![0-9])^[_a-zA-Z0-9]+$");

            try
            {
                new Program().Start(args);
            }
            catch (Exception e)
            {
                Console.WriteLine("\n\nError!\n" + e.ToString() + "\n");
                return;
            }

            if (autoExit == 0)
            {
                Console.WriteLine($"\n按任意键退出");
                Console.ReadLine();
            }
        }

        string sourceDirectory = AppContext.BaseDirectory;
        string targetDirectory = AppContext.BaseDirectory;
        string program = string.Empty;
        int searchOption = 1;
        int jsonformat = 0;
        string jsonGroup = string.Empty;
        string unGroupDirectory = string.Empty;
        bool unGroupOut;
        string sheetName = null;
        string jsonType = "JArray";
        string mainKey = string.Empty;
        string extensions = string.Empty;
        string startCell = "A1";
        string endCell = string.Empty;

        StringBuilder sb_temp = new StringBuilder();

        SearchOption SearchOption => searchOption == 1 ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

        public void Start(string[] args)
        {
            if (args != null) InputConfig(args);

            //program = "UniqueCharacter";
            //searchOption = 1;
            //extensions = ".txt,.meta,.cs";
            //sourceDirectory = @"E:\Code\Unity\Worker\2022\escape_urp_2024.02.04\Escpace\Assets\ResAll\xml\language\font\WordBank";
            //targetDirectory = @"D:\C07_Code\CShape\ExcelToConfigGame\ExcelToConfigGame\bin\Debug\net8.0\signalchars.txt";
            //jsonformat = -1;
            //jsonGroup = "all";
            
            switch (program)
            {
                case "MultFileLanguage":
                    if (File.Exists(sourceDirectory))
                    {
                        ExcelToMultFileLanguage(sourceDirectory);
                    }
                    break;
                case "Json":
                    unGroupOut = !string.IsNullOrEmpty(jsonGroup) && Directory.Exists(unGroupDirectory);

                    if (File.Exists(sourceDirectory))
                    {
                        ExcelToJson(sourceDirectory);
                    }
                    else if (Directory.Exists(sourceDirectory))
                    {
                        ExcelToJsonFromDir(sourceDirectory);
                    }
                    break;
                case "UniqueCharacter":
                    FileToUniqueCharacterFile(sourceDirectory);
                    break;
                default:
                    break;
            }
        }

        #region Excel to Json

        void ExcelToJsonFromDir(string sourceDirectory_) {

            string[] filePaths = Directory.GetFiles(sourceDirectory_, "*.xlsx", SearchOption);

            FileCount = filePaths.Length;
            FileNow = 0;

            Console.WriteLine($"\nExcelToJson [{FileNow},{FileCount}]:");

            for (int i = 0; i < filePaths.Length; i++) ExcelToJson(filePaths[i]);
        }

        void ExcelToJson(string filePath)
        {
            if (FileCount == 0) FileCount = 1;

            dynamic[] dataList;

            if (string.IsNullOrEmpty(endCell))
            {
                dataList = MiniExcel.Query(filePath, sheetName: sheetName, startCell: startCell.ToUpper(), configuration: new OpenXmlConfiguration() { FillMergedCells = true }).ToArray();
            }
            else
            {
                dataList = MiniExcel.QueryRange(filePath, sheetName: sheetName, startCell: startCell.ToUpper(), endCell:endCell.ToUpper(), configuration: new OpenXmlConfiguration() { FillMergedCells = true }).ToArray();
            }
            //Console.WriteLine(filePath);

            List<string[]> keywords = new List<string[]>();
            List<string> types = new List<string>();
            List<string> groups = new List<string>();

            string[] keywordValue = [];

            bool[] keywordValid = [];
            bool[] groupsValid = [];

            StringBuilder sb = new StringBuilder();

            bool isJArray = jsonType.Equals("JArray");
            int mainKeyIndex = -1;

            sb.Append(isJArray ? '[' : '{');
            bool one = false;
            bool uintext = true;

            foreach (var data in dataList)
            {
                var item = (data as IDictionary<string, object>)?.Values.ToArray();

                if (item == null || item.Length == 0) continue;

                if (item[0] != null && ((string)item[0]).StartsWith('#'))
                {
                    switch ((string)item[0])
                    {
                        case "#keyword":
                            keywords.Add(new string[item.Length-1]);

                            for (int i = 1; i < item.Length; i++)
                            {
                                keywords[keywords.Count-1][i-1] = (string)item[i];

                                if (!isJArray && mainKeyIndex == -1 && (((string)item[i]).Equals(mainKey))) 
                                {
                                    mainKeyIndex = i;
                                }
                            }
                            if (!isJArray && mainKeyIndex < 0)
                            {
                                Console.WriteLine($"Error JsonType is JArray but not find mainkey \"{mainKey}\"");
                                return;
                            }
                            break;
                        case "#type":
                            types.Clear();
                            for (int i = 1; i < item.Length; i++)
                            {
                                types.Add((item[i] == null ? string.Empty :(string)item[i]).ToLower());
                            }
                            break;
                        case "#group":
                            groups.Clear();
                            for (int i = 1; i < item.Length; i++)
                            {
                                groups.Add((item[i] == null ? string.Empty : (string)item[i]).ToLower());
                            }
                            break;
                        default:
                            break;
                    }
                    continue;
                }

                if (keywords.Count == 0) { Console.WriteLine($"#keyword 未被发现 Path:{filePath}"); return; }

                if (one) sb.Append(',');

                if (uintext) 
                {
                    uintext = false;
                    KeyWord(keywords, out keywordValue, out keywordValid);
                    Group(groups, out groupsValid);
                    Type(types, keywords[0].Length);

                    if (groups.Count == 0 && unGroupOut)
                    {
                        groupsValid = new bool[keywordValue.Length];
                        for (int i = 0; i < groupsValid.Length; i++) groupsValid[i] = true;
                    }

                    if (!string.IsNullOrEmpty(jsonGroup) && (groupsValid.Length == 0 || groupsValid.All(x=>!x))) {
                        
                        Console.WriteLine($"---- [{++FileNow},{FileCount}]");
                        return;
                    }
                }

                if (!isJArray)
                {
                    sb.Append($"\"{item[mainKeyIndex]}\":");
                }

                sb.Append('{');

                bool appedlined = false;
                for (int i = 1; i < item.Length; i++)
                {
                    if (!keywordValid[i - 1] ) continue;
                    if (groupsValid.Length > 0 && !groupsValid[i - 1]) continue;

                    string value = keywordValue[i - 1].Replace("*", ToValue(item[i], types[i - 1]));

                    if(appedlined) sb.Append(',');

                    sb.Append(value);
                    appedlined = true;

                }
                sb.Append('}');
                one = true;
            }

            sb.Append(isJArray ? ']' : '}');

            string jsontxt = sb.ToString();

            if (jsonformat >= 0)
            {
                if (isJArray)
                {
                    jsontxt = JArray.Parse(jsontxt).ToString(jsonformat == 0 ? Newtonsoft.Json.Formatting.None : Newtonsoft.Json.Formatting.Indented);
                }
                else
                {
                    jsontxt = JObject.Parse(jsontxt).ToString(jsonformat == 0 ? Newtonsoft.Json.Formatting.None : Newtonsoft.Json.Formatting.Indented);
                }
            }

            string fileName = Path.HasExtension(targetDirectory) ? Path.GetFileNameWithoutExtension(targetDirectory) : Path.GetFileNameWithoutExtension(filePath);
            string filePathO = Path.HasExtension(targetDirectory) ? targetDirectory : Path.Combine(targetDirectory, fileName + ".json");

            File.WriteAllText(filePathO, jsontxt);

            Console.WriteLine($"转换成功 [{++FileNow},{FileCount}] {Path.GetRelativePath(AppContext.BaseDirectory, filePath)} => {Path.GetRelativePath(AppContext.BaseDirectory, filePathO)}");
        }

        void Type(List<string> types,int Length) {

            if (types.Count == 0)
            {
                for (int i = 0; i < Length; i++) types.Add("string");
            }
        }

        void Group(List<string> groups, out bool[] result) {

            result = new bool[groups.Count];

            if (string.IsNullOrEmpty(jsonGroup) || jsonGroup.Equals("all"))
            {
                for (int i = 0; i < result.Length; i++) result[i] = true;
            }
            else
            {
                for (int i = 0; i < result.Length; i++)
                {
                    result[i] = groups[i].Equals("all") || groups[i].Equals(jsonGroup);
                }
            }
        }

        void KeyWord(List<string[]> keywords, out string[] line, out bool[] result)
        {
            line = new string[keywords[0].Length];
            result = new bool[keywords[0].Length];

            for (int i = 0; i < line.Length; i++)
            {
                string val = keywords[keywords.Count - 1][i];

                line[i] = string.IsNullOrEmpty(val) || !regex.IsMatch(val) ? string.Empty : $"\"{keywords[keywords.Count - 1][i]}\":*";
            }

            for (int i = keywords.Count - 1; i > 0; i--)
            {
                for (int j = 0; j < keywords[i].Length; j++)
                {
                    if (string.IsNullOrEmpty(keywords[i][j])) continue;

                    if (Equals(keywords[i][j], keywords[i - 1][j])) continue;               //本值 == 上值
                    else
                    {
                        bool l = j - 1 >= 0;                        //左值 有效
                        bool r = j + 1 < keywords[i].Length;        //右值 有效

                        bool ul = l && (!Equals(keywords[i - 1][j], keywords[i - 1][j - 1]));     //上值 != 上左值
                        bool ur = r && (!Equals(keywords[i - 1][j], keywords[i - 1][j + 1]));     //上值 != 上右值

                        if (ul || !l) line[j] = $"\"{keywords[i - 1][j]}\":" + "{" + line[j];
                        if (ur || !r) line[j] = line[j] + "}";
                    }
                }
            }

            for (int i = 0; i < result.Length; i++)
            {
                result[i] = !string.IsNullOrEmpty(line[i]);
            }
        }

        Dictionary<string, string[]> type2F = new Dictionary<string, string[]>();

        string ToValue(object obj, string type)
        {
            if (!type.Contains(':'))
            {
                switch (type)
                {
                    case "int":
                    case "float":
                    case "double":
                    case "long":
                    case "short":
                    case "byte":
                    case "bool":
                        return $"{(obj == null ? 0 : obj)}";
                    case "object":
                    case "json":
                        return $"{(obj == null ? "{}" : obj)}";
                    case "array":
                        return $"[{(obj == null ? string.Empty : obj.ToString().Replace(';',','))}]";
                    default:
                        return $"\"{obj}\"";
                }
            }
            else
            {
                if (type.StartsWith("array:"))
                {
                    if (obj == null || string.IsNullOrEmpty((string)obj)) return "[]";

                    sb_temp.Clear();

                    if (!type2F.ContainsKey(type)) type2F.Add(type, TypeSplit(type, "array:"));
                    var fg = type2F[type];

                    var items = ((string)obj).Split(';',StringSplitOptions.RemoveEmptyEntries);

                    sb_temp.Append('[');

                    for (int i = 0; i < items.Length; i++)
                    { 
                        string item = items[i];

                        sb_temp.Append('{');
                        var ite = item.Remove(item.Length - 1, 1).Remove(0, 1).Split(',', StringSplitOptions.RemoveEmptyEntries);

                        for (var j = 0; j < ite.Length; j++)
                        {
                            sb_temp.Append('"');
                            sb_temp.Append(fg[j * 2 + 0]);
                            sb_temp.Append('"');
                            sb_temp.Append(':');
                            sb_temp.Append(ToValue(ite[j], fg[j * 2 + 1]));
                            if (j < ite.Length - 1) sb_temp.Append(',');
                        }

                        sb_temp.Append('}');
                        if (i < items.Length - 1) sb_temp.Append(',');
                    }
                    sb_temp.Append(']');
                    return ToValue(sb_temp.ToString(), "object");
                }
                else if (type.StartsWith("object:"))
                {
                    if (obj == null || string.IsNullOrEmpty((string)obj)) return "{}";

                    sb_temp.Clear();

                    if (!type2F.ContainsKey(type)) type2F.Add(type, TypeSplit(type, "object:"));
                    var fg = type2F[type];

                    sb_temp.Append('{');
                    var ite = ((string)obj).Split(',', StringSplitOptions.RemoveEmptyEntries);

                    for (var j = 0; j < ite.Length; j++)
                    {
                        sb_temp.Append('"');
                        sb_temp.Append(fg[j * 2 + 0]);
                        sb_temp.Append('"');
                        sb_temp.Append(':');
                        sb_temp.Append(ToValue(ite[j], fg[j * 2 + 1]));
                        if (j < ite.Length - 1) sb_temp.Append(',');
                    }

                    sb_temp.Append('}');
                    return ToValue(sb_temp.ToString(), "object");
                }

                return $"\"{obj}\"";
            }
        }

        #endregion

        #region Excel to MultFileLanguage

        void ExcelToMultFileLanguage(string filePath)
        {
            //string filename = Path.GetFileNameWithoutExtension(filePath);
            //string nfilePath = Path.Combine(targetDirectory, filename + ".txt");

            var dataList = MiniExcel.Query(filePath,sheetName: sheetName).ToArray();

            List<string> keywords = new List<string>();
            List<StringBuilder> texts = new List<StringBuilder>();

            foreach (var data in dataList)
            {
                var item = (data as IDictionary<string, object>)?.Values.ToArray();

                if (item == null || item.Length == 0) continue;
                
                if (item[0] != null && ((string)item[0]).StartsWith('#'))
                {
                    if (item.Length > 1 && item[1] != null && item[1].Equals("keyword"))
                    {
                        for (int i = 2; i < item.Length; i++)
                        {
                            if (texts.Count < i - 1)
                            {
                                texts.Add(new StringBuilder());
                                keywords.Add(item[i] == null ? string.Empty : (string)item[i]);
                            }
                        }
                    }
                    continue;
                }

                if (keywords.Count == 0) { throw new Exception($"#keyword 未被发现 Path:{filePath}");  return; }

                for (int i = 2; i < item.Length; i++)
                {
                    texts[i - 2].AppendLine(item[1] + "\t" + item[i]);
                }
            }

            for (int i = 0; i < texts.Count; i++)
            {
                if (string.IsNullOrEmpty(keywords[i])) { ++FileNow; continue; }

                string fileName = keywords[i];
                string filePathO = Path.Combine(targetDirectory, fileName + ".txt");
                File.WriteAllText(filePathO, texts[i].ToString());

                Console.WriteLine($"转换成功 - 分支 [{++FileNow},{FileCount} - {i+1}] {Path.GetRelativePath(AppContext.BaseDirectory, filePath)} => {Path.GetRelativePath(AppContext.BaseDirectory, filePathO)}");
            }
        }

        void ExcelToTabTxt(string filePath)
        {
            string filename = Path.GetFileNameWithoutExtension(filePath);
            string nfilePath = Path.Combine(targetDirectory, filename + ".txt");

            FileStream stream = File.OpenRead(filePath);

            var dataList = stream.Query().ToArray();

            StringBuilder sb = new StringBuilder();

            foreach (var data in dataList)
            {
                var item = (data as IDictionary<string, object>)?.Values;

                string ms = string.Empty;

                foreach (var key in item)
                {
                    ms += $"{key}\t";
                }

                sb.AppendLine(ms);
            }

            File.WriteAllText(nfilePath, sb.ToString());

            Console.WriteLine($"转换成功 {filename}");
        }

        #endregion

        #region MulFile to UniqueCharacter File

        void FileToUniqueCharacterFile(string filePaths_) {

            string[] paths = filePaths_.Split("::", StringSplitOptions.RemoveEmptyEntries);

            List<string> filePaths = new List<string>();

            List<string> extension = new List<string>() { 
                ".txt",
                ".xml",
                ".json",
                ".yml"
            };

            if (!string.IsNullOrEmpty(extensions))
            {
                extension.Clear();

                var exts = extensions.Split(',',StringSplitOptions.RemoveEmptyEntries);

                foreach (var ext in exts) { extension.Add(ext); }
            }

            foreach (var item in paths)
            {
                if (Directory.Exists(item))
                {
                    var files = Directory.GetFiles(item,"*.*", SearchOption);

                    for (int i = 0; i < files.Length; i++)
                    {
                        if (!extension.Contains(Path.GetExtension(files[i]))) continue;

                        filePaths.Add(files[i]);
                    }
                }
                else if (File.Exists(item))
                {
                    if (!extension.Contains(Path.GetExtension(item))) continue;

                    filePaths.Add(item);
                }
            }

            List<char> chars = new List<char>();

            StringBuilder sb = new StringBuilder();

            Console.WriteLine("找到文件"+filePaths.Count);

            for (int n = 0; n < filePaths.Count; n++)
            {
                var text = File.ReadAllText(filePaths[n]);
                int count = 0;
                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];

                    if (!chars.Contains(c))
                    {
                        chars.Add(c);
                        count++;
                    }
                }

                Console.WriteLine($"已读取文件:[{n + 1},{filePaths.Count}] -[char,{count}] {Path.GetFileName(filePaths[n])}");
            }

            chars.Sort();

            for (int i = 0; i < chars.Count; i++)
            {
                sb.Append(chars[i]);
            }

            File.WriteAllText(targetDirectory, sb.ToString(),Encoding.UTF8);

            Console.WriteLine($"已将所有字符写入文件 [char,{sb.Length}]" + targetDirectory);
        }
        #endregion
        void InputConfig(string[] args)
        {
            var list = args.ToList();

            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].Equals("sourceDirectory"))
                {
                    if ((i + 1) < list.Count)
                    {
                        sourceDirectory = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                if (list[i].Equals("targetDirectory"))
                {
                    if ((i + 1) < list.Count)
                    {
                        targetDirectory = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
                
                if (list[i].Equals("searchOption"))
                {
                    if ((i + 1) < list.Count)
                    {
                        if (int.TryParse(list[i + 1], out int searchOption_))
                        {
                            searchOption = searchOption_;
                        }
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                if (list[i].Equals("jsonformat"))
                {
                    if ((i + 1) < list.Count)
                    {
                        if (int.TryParse(list[i + 1], out int jsonformat_))
                        {
                            jsonformat = jsonformat_;
                        }
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
                
                if (list[i].Equals("program"))
                {
                    if ((i + 1) < list.Count)
                    {
                        program = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                if (list[i].Equals("autoExit"))
                {
                    if ((i + 1) < list.Count)
                    {
                        if (int.TryParse(list[i + 1], out int autoExit_))
                        {
                            Program.autoExit = autoExit_;
                        }
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                if (list[i].Equals("jsonGroup"))
                {
                    if ((i + 1) < list.Count)
                    {
                        jsonGroup = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
                
                if (list[i].Equals("unGroupDirectory"))
                {
                    if ((i + 1) < list.Count)
                    {
                        unGroupDirectory = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                if (list[i].Equals("sheetName"))
                {
                    if ((i + 1) < list.Count)
                    {
                        sheetName = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                if (list[i].Equals("jsonType"))
                {
                    if ((i + 1) < list.Count)
                    {
                        jsonType = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                if (list[i].Equals("mainkey"))
                {
                    if ((i + 1) < list.Count)
                    {
                        mainKey = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                if (list[i].Equals("extensions"))
                {
                    if ((i + 1) < list.Count)
                    {
                        extensions = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
                
                if (list[i].Equals("startCell"))
                {
                    if ((i + 1) < list.Count)
                    {
                        startCell = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                if (list[i].Equals("endCell"))
                {
                    if ((i + 1) < list.Count)
                    {
                        endCell = list[i + 1];
                        list.RemoveAt(i + 1);
                        list.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
            }
        }
    }
}
