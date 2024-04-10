
using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;

/**
 * 通用方法
 * <p>
 * by miracleren@gmail.com
 */

namespace NiceDoc.Net
{

    public class NiceUtils
    {

        /**
         * {{par}} 参数查找正则
         *
         * @param str 查找串
         * @return 返结果
         */
        public static MatchCollection getMatchingLabels(String str)
        {
            string pattern = "(?<=\\{\\{)(.+?)(?=\\}\\})";
            MatchCollection matcher = Regex.Matches(str, pattern);
            return matcher;
        }

        /**
         * 补全label格式
         *
         * @param label
         * @return
         */
        public static String labelFormat(String label)
        {
            return "{{" + label + "}}";
        }

        /**
         * 实体类转map
         *
         * @param entity
         * @return
         */
        public static Dictionary<string, object> entityToDictionary(object entity)
        {
            Dictionary<string, object> map = new Dictionary<string, object>();
            foreach (PropertyInfo field in entity.GetType().GetProperties())
            {
                try
                {
                    map.Add(field.Name, field.GetValue(entity, null));
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            return map;
        }

        /**
         * 实体类列表转map列表
         *
         * @param entityList
         * @return
         */
        public static List<Dictionary<string, object>> listEntityToDictionary(List<object> entityList)
        {
            List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
            foreach (Object entity in entityList)
            {
                list.Add(entityToDictionary(entity));
            }
            return list;
        }

        /**
         * 转sting方法
         *
         * @param object
         * @return
         */
        public static string toString(object val)
        {
            return val == null ? "" : val.ToString();
        }

        /**
     * 判断对象是否是数值
     *
     * @param object
     * @return
     */
        public static bool isNumber(Object val)
        {
            Type type = val.GetType();
            return type.IsPrimitive && type != typeof(bool) && type != typeof(char);
        }
    }
}
