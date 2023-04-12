using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

/**
 * 项目：AutoOpen-PPT
 * 作者：20信息管理-YYH
 **/
namespace AutoOpen_PPT
{
    public partial class ThisAddIn
    {
        private static readonly DateTime Jan1st1970 = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        private const string openedName = "OpenedPres.txt";
        private readonly Dictionary<string, long> closed = new Dictionary<string, long>();
        private readonly List<string> closedPres = new List<string>();
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // 向字典加入特殊键值对，记录程序最终关闭时的时间；
            closed.Add("$Time$", CurrentTimeMillis());
            string openedTemp = Path.Combine(Path.GetTempPath(), openedName);
            File.WriteAllLines(openedTemp, Serialize(closed));
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var lastClosed = 0L;
            string openedTemp = Path.Combine(Path.GetTempPath(), openedName);
            if (File.Exists(openedTemp))
            {
                var map = new Dictionary<string, long>();
                var raw = Deserialize(File.ReadAllLines(openedTemp).ToList());
                foreach (var entry in raw)
                {
                    if (entry.Key == "$Time$")
                    {
                        lastClosed = entry.Value;
                        continue;
                    }
                    map.Add(entry.Key, entry.Value);
                }
                foreach (var entry in map)
                {
                    // 在程序最终关闭前30秒内关闭的项目才会被重新打开；
                    if (lastClosed - entry.Value < 30000L) Application.Presentations.Open(entry.Key);
                }
                File.Delete(openedTemp);
            }
            this.Application.PresentationClose += Application_PresentationClose;
        }

        private void Application_PresentationClose(PowerPoint.Presentation Pres)
        {
            closed.Add(Pres.FullName, CurrentTimeMillis());
        }

        // 获取协调世界时，转换为长整型；
        private long CurrentTimeMillis()
        {
            return (long)(DateTime.UtcNow - Jan1st1970).TotalMilliseconds;
        }

        // 将字典序列化为字符串列表，字符串的格式是"文件名>时间"；
        private List<string> Serialize(Dictionary<string, long> o)
        {
            var list = new List<string>();
            foreach (var e in o)
            {
                list.Add(e.Key + ">" + e.Value.ToString());
            }
            return list;
        }

        // 将字符串列表反序列化为字典；
        private Dictionary<string, long> Deserialize(List<string> list)
        {
            var dic = new Dictionary<string, long>();
            foreach (var o in list)
            {
                if (!o.Contains('>')) continue;
                var splitted = o.Split('>');
                var path = splitted[0];
                long.TryParse(splitted[1], out long time);
                if (time == 0L) continue;
                dic.Add(path, time);
            }
            return dic;
        }
    }
}
