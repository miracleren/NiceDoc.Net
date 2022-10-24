using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NiceDoc.Net
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(" _   _ _          _____             ");
            Console.WriteLine("| \\ | (_)        |  __ \\            ");
            Console.WriteLine("|  \\| |_  ___ ___| |  | | ___   ___ ");
            Console.WriteLine("| . ` | |/ __/ _ \\ |  | |/ _ \\ / __|");
            Console.WriteLine("| |\\  | | (_|  __/ |__| | (_) | (__ ");
            Console.WriteLine("|_| \\_|_|\\___\\___|_____/ \\___/ \\___|");

            //测试示例模板
            string path = Environment.CurrentDirectory + "/doc/";

            NiceDoc docx = new NiceDoc(path + "test.docx");

            Dictionary<string, object> labels = new Dictionary<string, object>();
            //值标签
            labels.Add("startTime", "1881年9月25日");
            labels.Add("endTime", "1936年10月19日");
            labels.Add("title", "精选作品目录");
            labels.Add("press", "鲁迅同学出版社");

            //枚举标签
            labels.Add("likeBook", 2);
            //布尔标签
            labels.Add("isQ", true);
            //等于
            labels.Add("isNew", 2);
            //多选二进制值
            labels.Add("look", 3);
            //if语句
            labels.Add("showContent", 2);
            //日期格式标签
            labels.Add("printDate", DateTime.Now);

            docx.pushLabels(labels);

            //表格
            List<Dictionary<string, object>> books = new List<Dictionary<string, object>>();
            Dictionary<string, object> book1 = new Dictionary<string, object>();
            book1.Add("name", "汉文学史纲要");
            book1.Add("time", "1938年，鲁迅全集出版社");
            books.Add(book1);
            Dictionary<string, object> book2 = new Dictionary<string, object>();
            book2.Add("name", "中国小说史略");
            book2.Add("time", "1923年12月，上册；1924年6月，下册");
            books.Add(book2);
            docx.pushTable("books", books);

            //生成文档
            docx.save(path, Guid.NewGuid() + ".docx");
        }
    }
}
