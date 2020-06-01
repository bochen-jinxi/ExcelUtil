﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace ExcelUtil
{
    public class Base
    {
        public Base()
        {
            //Encoding.UTF8
            Encoding.GetEncoding(65001);
        }

        /// <summary>
        ///     获取根目录
        /// </summary>
        /// <returns></returns>
        public string GetTestRootPath()
        {
            return Directory.GetCurrentDirectory();
        }

        /// <summary>
        ///     获取文件路径
        /// </summary>
        /// <param name="paths"></param>
        /// <returns></returns>
        public string GetTestFilePath(params string[] paths)
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
        ///     删除文件
        /// </summary>
        public void DeleteFile(string path)
        {
            if (File.Exists(path)) File.Delete(path);
        }

        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>文件存在性</returns>
        public static bool TryDelete(string fileName)
        {
            var index = 1;
            var exist = File.Exists(fileName);
            var ts = 1;

            while (exist && index < 6)
            {
                try
                {
                    File.Delete(fileName);
                }
                finally
                {
                    exist = File.Exists(fileName);
                    if (exist)
                    {
                        Thread.Sleep(ts * 10);
                        ts *= index++;
                    }
                }
            }
            return exist;
        }

        #region 数据处理

        /// <summary>
        /// ByteToMemoryStream
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public static MemoryStream B2MS(byte[] buffer)
        {
            return new MemoryStream(buffer);
        }

        /// <summary>
        /// JsonToData
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="str"></param>
        /// <returns></returns>
        public static T GetData<T>(string str)
        {
            return JsonConvert.DeserializeObject<T>(str);
        }

        /// <summary>
        /// DataToJson
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string GetJson(object obj)
        {
            return JsonConvert.SerializeObject(obj);
        }

        #endregion 数据处理
    }
}