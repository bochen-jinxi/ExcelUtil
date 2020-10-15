using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelUtil._05_Utils
{
    /// <summary>
    /// 文件工具
    /// </summary>
   public static class FileUtil
    {
        /// <summary>
        /// 获取根目录
        /// </summary>
        /// <returns></returns>
        public static string GetTestRootPath()
        {
            return Directory.GetCurrentDirectory();
        }

        /// <summary>
        /// 获取文件路径
        /// </summary>
        /// <param name="paths"></param>
        /// <returns></returns>
        public static string GetTestFilePath(params string[] paths)
        {
            var rootPath = GetTestRootPath();
            var list = new List<string>
            {
                rootPath
            };
            list.AddRange(paths);
            return Path.Combine(list.ToArray());
        }

        /// <summary>
        /// 删除文件
        /// </summary>
        public static void DeleteFile(string path)
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
