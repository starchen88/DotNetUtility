using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace DotNetUtility
{
    /// <summary>
    /// Office操作类
    /// </summary>
    public static class OfficeHelper
    {
        #region 导出数据到Excel
        /// <summary>
        /// 基于表达式规则将IEnumerable<typeparamref name="T"/>转换为DataTable
        /// 请注意，这没有使用反射之类的自动化方式，需要你自行提供转换策略
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="mapExp">表示转换策略的表达式字典，key为列名，value为获取值的Lambda表达式</param>
        /// <returns></returns>
        public static DataTable ToDatatable<T>(IEnumerable<T> source, Dictionary<string, Expression<Func<T, object>>> mapExp)
        {
            DataTable dt = new DataTable();
            foreach (var item in mapExp)
            {
                Type t = typeof(object);
                if (item.Value.Body.NodeType == ExpressionType.MemberAccess)
                {
                    var member = ((MemberExpression)item.Value.Body).Member;
                    if (member is PropertyInfo)
                    {
                        t = ((PropertyInfo)member).PropertyType;
                    }
                    else
                    {
                        throw new NotSupportedException();
                    }
                }
                else if (item.Value.Body.NodeType == ExpressionType.Convert)
                {
                    var u = (UnaryExpression)item.Value.Body;
                    t = u.Operand.Type;
                }
                else if (item.Value.Body.NodeType == ExpressionType.Call)
                {
                    var mc = (MethodCallExpression)item.Value.Body;
                    t = mc.Type;
                }
                else
                {
                    t = typeof(object);
                }
                if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    t = Nullable.GetUnderlyingType(t);
                }
                dt.Columns.Add(item.Key, t);
            }
            var mapFunc = mapExp.ToDictionary(it => it.Key, it => it.Value.Compile());
            foreach (var dto in source)
            {
                var row = dt.NewRow();
                foreach (var map in mapFunc)
                {
                    row[map.Key] = map.Value(dto) ?? DBNull.Value;
                }
                dt.Rows.Add(row);
            }
            return dt;
        }
        #endregion
        #region 导入Excel
        public static string ErrorFormatTRCM = "数据或格式错误：表:{0} 行:{1} 列:{2},{3}";
        public static string ErrorFormatTRM = "数据或格式错误：表:{0} 行:{1},{2}";
        /// <summary>
        /// 从DataTable导入数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="table"></param>
        /// <param name="convertFuncDic">包含各字段映射规则的字典，key为字段名，value为转换逻辑，返回true表示转换成功、false表示转换失败</param>
        /// <returns>返回值包含转换成功/失败状态，转换失败的消息和转换成功的数据</returns>
        public static IEnumerable<ConvertFromDatatableResult<T>> ConvertFromDatatable<T>(DataTable table, Dictionary<string, Func<T, object, DataRow, bool>> convertFuncDic) where T : new()
        {
            for (int r = 0; r < table.Rows.Count; r++)
            {
                DataRow row = table.Rows[r];
                var item = new T();
                var result = new ConvertFromDatatableResult<T>();
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    DataColumn column = table.Columns[c];
                    if (convertFuncDic.ContainsKey(column.ColumnName))
                    {
                        try
                        {
                            var val = row[column];
                            if (val == null || val == DBNull.Value)
                            {//为空，跳过

                            }
                            else
                            {
                                if (convertFuncDic[column.ColumnName](item, val, row))
                                {
                                    result.IsSuccess = true;                                    
                                }
                                else
                                {
                                    result.IsSuccess = false;
                                    result.Message = string.Format(ErrorFormatTRCM, table.TableName, r + 2, c + 1, column.ColumnName + "格式不正确");
                                }
                            }
                        }
                        catch (Exception)
                        {
                            result.IsSuccess = false;
                            result.Message = string.Format(ErrorFormatTRCM, table.TableName, r + 2, c + 1, column.ColumnName + "格式不正确");
                        }
                    }
                    else
                    {//跳过
                        continue;
                    }
                }
                //调用数据验证
                var validationContext = new ValidationContext(item);
                var validationResults = new List<ValidationResult>();
                Validator.TryValidateObject(item, validationContext, validationResults, false);
                foreach (var vr in validationResults)
                {
                    result.IsSuccess = false;
                    result.Message = string.Format(ErrorFormatTRM, table.TableName, r + 2, vr.ErrorMessage);
                }
                if (result.IsSuccess)
                {
                    result.Data = item;
                }
                yield return result;
            }
        }
        public static bool TryConvertThen<T>(object v, Action<T> action)
        {
            var t = typeof(T);
            var vt = v.GetType();
            if (t == vt)
            {
                action((T)v);
                return true;
            }
            else
            {
                var converter = TypeDescriptor.GetConverter(t);
                if (converter.CanConvertFrom(vt))
                {
                    var r = (T)converter.ConvertFrom(v);
                    action(r);
                    return true;
                }
                else
                {
                    T r = (T)Convert.ChangeType(v, t);
                    action(r);
                    return true;
                }
            }
        }        
        #endregion
    }

    public class ConvertFromDatatableResult<T>
    {
        /// <summary>
        /// 转换成功的数据
        /// </summary>
        public T Data { get; set; }
        /// <summary>
        /// 转换成功
        /// </summary>
        public bool IsSuccess { get; set; }
        /// <summary>
        /// 转换失败的消息提示
        /// </summary>
        public string Message { get; set; }
    }
}
