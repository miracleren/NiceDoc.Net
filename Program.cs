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


            //测试示例模板生成word
            TestTemplate.buildTestDocx();

            //测试示例模板生成xlsx
            TestTemplate.buildTestXlsx();
        }
    }
}
