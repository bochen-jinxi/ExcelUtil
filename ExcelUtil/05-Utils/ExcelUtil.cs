﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;

namespace ExcelUtil
{
    /// <summary>
    /// 三方源码工具
    /// </summary>
    public static class ExcelUtil
    {
        /// <summary>
        /// 将集合转成DataTable
        /// </summary>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(IEnumerable<T> source)
        {
            var table = Cache<T>.SchemeFactory();

            foreach (var item in source)
            {
                var row = table.NewRow();
                Cache<T>.Fill(row, item);
                table.Rows.Add(row);
            }
            return table;
        }

        private static class Cache<T>
        {
            // ReSharper disable StaticMemberInGenericType
            private static readonly PropertyInfo[] PropertyInfos;
            // ReSharper restore StaticMemberInGenericType

            // ReSharper disable StaticMemberInGenericType
            public static readonly Func<DataTable> SchemeFactory;
            // ReSharper restore StaticMemberInGenericType
            public static readonly Action<DataRow, T> Fill;

            static Cache()
            {
                PropertyInfos = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                SchemeFactory = GenerateSchemeFactory();
                Fill = GenerateFill();
            }

            private static Func<DataTable> GenerateSchemeFactory()
            {
                var dynamicMethod = new DynamicMethod($"__Extensions__SchemeFactory__Of__{typeof(T).Name}", typeof(DataTable), Type.EmptyTypes, typeof(T), true);

                var generator = dynamicMethod.GetILGenerator();
                // ReSharper disable AssignNullToNotNullAttribute
                generator.Emit(OpCodes.Newobj, typeof(DataTable).GetConstructor(Type.EmptyTypes));
                // ReSharper restore AssignNullToNotNullAttribute
                generator.Emit(OpCodes.Dup);

                // ReSharper disable PossibleNullReferenceException
                generator.Emit(OpCodes.Callvirt, typeof(DataTable).GetProperty("Columns").GetMethod);
                // ReSharper restore PossibleNullReferenceException

                foreach (var propertyInfo in PropertyInfos)
                {
                    generator.Emit(OpCodes.Dup);
                    generator.Emit(OpCodes.Ldstr, propertyInfo.Name);
                    generator.Emit(OpCodes.Ldtoken, (propertyInfo.PropertyType.IsGenericType) && (propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>)) ? propertyInfo.PropertyType.GetGenericArguments()[0] : propertyInfo.PropertyType);
                    // ReSharper disable AssignNullToNotNullAttribute
                    generator.Emit(OpCodes.Call,
                        typeof(Type).GetMethod("GetTypeFromHandle", new[] { typeof(RuntimeTypeHandle) }));
                    // ReSharper restore AssignNullToNotNullAttribute
                    // ReSharper disable AssignNullToNotNullAttribute
                    generator.Emit(OpCodes.Callvirt,
                        typeof(DataColumnCollection).GetMethod("Add", new[] { typeof(string), typeof(Type) }));
                    // ReSharper restore AssignNullToNotNullAttribute
                    generator.Emit(OpCodes.Pop);
                }

                generator.Emit(OpCodes.Pop);
                generator.Emit(OpCodes.Ret);

                return (Func<DataTable>)dynamicMethod.CreateDelegate(typeof(Func<DataTable>));
            }

            private static Action<DataRow, T> GenerateFill()
            {
                var dynamicMethod =
                    new DynamicMethod($"__Extensions__Fill__Of__{typeof(T).Name}", typeof(void),
                        new[] { typeof(DataRow), typeof(T) }, typeof(T), true);

                var generator = dynamicMethod.GetILGenerator();
                for (var i = 0; i < PropertyInfos.Length; i++)
                {
                    generator.Emit(OpCodes.Ldarg_0);
                    switch (i)
                    {
                        case 0:
                            generator.Emit(OpCodes.Ldc_I4_0);
                            break;
                        case 1:
                            generator.Emit(OpCodes.Ldc_I4_1);
                            break;
                        case 2:
                            generator.Emit(OpCodes.Ldc_I4_2);
                            break;
                        case 3:
                            generator.Emit(OpCodes.Ldc_I4_3);
                            break;
                        case 4:
                            generator.Emit(OpCodes.Ldc_I4_4);
                            break;
                        case 5:
                            generator.Emit(OpCodes.Ldc_I4_5);
                            break;
                        case 6:
                            generator.Emit(OpCodes.Ldc_I4_6);
                            break;
                        case 7:
                            generator.Emit(OpCodes.Ldc_I4_7);
                            break;
                        case 8:
                            generator.Emit(OpCodes.Ldc_I4_8);
                            break;
                        default:
                            if (i <= 127)
                            {
                                generator.Emit(OpCodes.Ldc_I4_S, (byte)i);
                            }
                            else
                            {
                                generator.Emit(OpCodes.Ldc_I4, i);
                            }
                            break;
                    }
                    generator.Emit(OpCodes.Ldarg_1);
                    generator.Emit(OpCodes.Callvirt, PropertyInfos[i].GetGetMethod(true));
                    if (PropertyInfos[i].PropertyType.IsValueType)
                    {
                        generator.Emit(OpCodes.Box, PropertyInfos[i].PropertyType);
                    }
                    // ReSharper disable AssignNullToNotNullAttribute
                    generator.Emit(OpCodes.Callvirt, typeof(DataRow).GetMethod("set_Item", new[] { typeof(int), typeof(object) }));
                    // ReSharper restore AssignNullToNotNullAttribute
                }
                generator.Emit(OpCodes.Ret);
                return (Action<DataRow, T>)dynamicMethod.CreateDelegate(typeof(Action<DataRow, T>));
            }

        }

        /// <summary>
        ///     将集合转成DataTable
        /// </summary>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(ICollection<T> source)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p =>
                new DataColumn(p.Name,
                    p.PropertyType.IsGenericType && p.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>)
                        ? p.PropertyType.GetGenericArguments()[0]
                        : p.PropertyType)).ToArray());

            for (var i = 0; i < source.Count; i++)
            {
                var tempList = new ArrayList();
                foreach (var obj in props.Select(pi => pi.GetValue(source.ElementAt(i), null))) tempList.Add(obj);
                var array = tempList.ToArray();
                dt.LoadDataRow(array, true);
            }

            return dt;
        }
        /// <summary>
        ///     将DataTable转List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static IList<T> ToList<T>(DataTable dt) where T : class
        {
            IList<T> list = new List<T>();
            string tempName = "";
            foreach (DataRow dr in dt.Rows)
            {
                T t = Activator.CreateInstance<T>();
                var props = typeof(T).GetProperties();
                foreach (var pro in props)
                {
                    tempName = pro.Name;
                    if (!dt.Columns.Contains(tempName)) continue;
                    if (!pro.CanWrite) continue;
                    var value = dr[tempName];
                    if (value != DBNull.Value)
                        pro.SetValue(t, value, null);
                }

                list.Add(t);
            }

            return list;
        }
    }
}