using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NiceDoc.Net
{
    public static class TestTemplate
    {
        static string path = Environment.CurrentDirectory + "/doc/";

        /**
        * 测试示例模板生成word
        */
        public static void buildTestDocx()
        {
            //测试示例模板


            NiceDoc docx = new NiceDoc(path + "test.docx");

            Dictionary<string, object> labels = new Dictionary<string, object>();
            //值标签
            labels.Add("startTime", "1881年9月25日");
            labels.Add("endTime", "1936年10月19日");
            labels.Add("title", "精选作品目录");
            labels.Add("press", "鲁迅同学出版社");

            //枚举标签
            labels.Add("likeBook", 3);
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

            //添加头像
            labels.Add("headImg", path + "head.png");

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


        /**
     * 测试示例模板生成xlsx
     */
        public static void buildTestXlsx()
        {
            //测试示例模板生成word
            NiceExcel excel = new NiceExcel(path + "test.xlsx");

            Dictionary<string, object> labels = new Dictionary<string, object>();
            //值标签
            labels.Add("date", "2023年1月1日");
            labels.Add("title", "精选作品统计");
            //枚举标签
            labels.Add("likeBook", 2);
            //多选二进制值
            labels.Add("lookType", 3);
            //if语句
            labels.Add("showBanner", 1);
            //日期格式标签
            labels.Add("printDate", DateTime.Now);

            excel.pushLabels(labels);


            //表格
            List<Dictionary<string, object>> books = new List<Dictionary<string, object>>();
            for (int i = 0; i <= 10; i++)
            {
                Dictionary<string, object> book = new Dictionary<string, object>();
                book.Add("name", "汉文学史纲要" + i);
                book.Add("time", 1900 + i + "年");
                book.Add("intro", "简明扼要的介绍，本书是一本好书，推荐" + i + "星");
                book.Add("byName", "作者" + i + "号");
                book.Add("pages", i * 100);
                books.Add(book);
            }
            excel.pushTable("books", books);


            //生成文档
            excel.save(path, Guid.NewGuid() + ".xlsx");
        }
    }
}
