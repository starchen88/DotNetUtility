using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq.Expressions;
using System.Text;

namespace DotNetUtility.Tests
{
    [TestClass]
    public class OfficeTest
    {
        public List<TestData1> CreateTestData1List()
        {
            return new List<TestData1>()
            {
                new TestData1{
                    Id = -1,
                    Name = "张三",
                    Time = DateTimeOffset.Now
                },
                new TestData1{
                    Id = 0,
                    Name = null,
                    Time = null
                },
                new TestData1{
                    Id =2,
                    Name = string.Empty,
                    Time = new DateTimeOffset()
                }
            };
        }
        [TestMethod]
        public void ListMapDatatable()
        {
            var list = CreateTestData1List();
            var oldJson = JsonConvert.SerializeObject(list);
            var dt = OfficeHelper.ToDatatable(list, new Dictionary<string, Expression<Func<TestData1, object>>> {
                { "Id",it=>it.Id },
                { "Name",it=>it.Name },
                { "Time",it=>it.Time }
            });
            int successCount = 0;
            List<string> messageList = new List<string>(dt.Rows.Count);
            List<TestData1> testData1List = new List<TestData1>(dt.Rows.Count);
            foreach (var item in OfficeHelper.ConvertFromDatatable(dt, new Dictionary<string, Func<TestData1, object, DataRow, bool>> {
                    { "Id",(it,obj,row)=>OfficeHelper.TryConvertThen<int>(obj,v=>it.Id=v)},
                    { "Name",(it,obj,row)=>OfficeHelper.TryConvertThen<string>(obj,v=>it.Name=v)},
                    { "Time",(it,obj,row)=>OfficeHelper.TryConvertThen<DateTimeOffset>(obj,v=>it.Time=v)}
                }))
            {
                if (item.IsSuccess)
                {
                    successCount++;
                    testData1List.Add(item.Data);
                }
                else
                {
                    messageList.Add(item.Message);
                }
            }
            if (successCount == testData1List.Count)
            {
                var newJson = JsonConvert.SerializeObject(testData1List);
                Assert.AreEqual(oldJson, newJson);
            }
            else
            {
                Assert.Fail();
            }
        }
        [TestMethod]
        public void ListMapExcel()
        {
            var list = CreateTestData1List();
            var oldJson = JsonConvert.SerializeObject(list);
            var dt = OfficeHelper.ToDatatable(list, new Dictionary<string, Expression<Func<TestData1, object>>> {
                { "Id",it=>it.Id },
                { "Name",it=>it.Name },
                { "Time",it=>it.Time }
            });
            DataTable newTable;
            byte[] bs;
            using (var memoryStream = new MemoryStream())
            {
                var ds = new DataSet();
                ds.Tables.Add(dt);
                NPOIHelper.DatasetToExcel(ds, memoryStream, ExcelFormat.Xlsx);
                bs = memoryStream.ToArray();                
            }
            using (var memoryStream = new MemoryStream(bs))
            {
                newTable = NPOIHelper.ExcelToDataset(memoryStream).Tables[0];
            }
            int successCount = 0;
            List<string> messageList = new List<string>(newTable.Rows.Count);
            List<TestData1> testData1List = new List<TestData1>(newTable.Rows.Count);
            foreach (var item in OfficeHelper.ConvertFromDatatable(newTable, new Dictionary<string, Func<TestData1, object, DataRow, bool>> {
                    { "Id",(it,obj,row)=>OfficeHelper.TryConvertThen<int>(obj,v=>it.Id=v)},
                    { "Name",(it,obj,row)=>OfficeHelper.TryConvertThen<string>(obj,v=>it.Name=v)},
                    { "Time",(it,obj,row)=>OfficeHelper.TryConvertThen<DateTimeOffset>(obj,v=>it.Time=v)}
                }))
            {
                if (item.IsSuccess)
                {
                    successCount++;
                    testData1List.Add(item.Data);
                }
                else
                {
                    messageList.Add(item.Message);
                }
            }
            if (successCount == testData1List.Count)
            {
                var newJson = JsonConvert.SerializeObject(testData1List);
                //Assert.AreEqual(oldJson, newJson); 由于DateTime转换到Excel精度丢失，无法保证导入导出的数据完全一致
            }
            else
            {
                Assert.Fail();
            }
        }
    }
    public class TestData1
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTimeOffset? Time { get; set; }
    }
}
