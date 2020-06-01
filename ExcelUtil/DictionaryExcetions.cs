using System.Collections.Generic;
using System.Linq;

namespace ExcelUtil
{
    /// <summary>
    /// 字典扩展
    /// </summary>
    public static class DictionaryExcetions
    {
        /// <summary>
        /// 字典转换
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="instance"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static List<T> Get<T>(this Dictionary<string, List<object>> instance, string name)
        {
            List<T> list = new List<T>();
            if (instance != null && instance.Any() && instance[name] != null && instance[name].Any())
            {
                foreach (var data in instance[name])
                {
                    list.Add((T)data);
                }
            }
            return list;
        }

        /// <summary>
        /// 字典转换
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="instance"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static T Get<T>(this Dictionary<string, object> instance, string name)
        {
            if (instance[name] != null)
            {
                return (T)instance[name];
            }
            else
            {
                return default(T);
            }
        }
    }
}