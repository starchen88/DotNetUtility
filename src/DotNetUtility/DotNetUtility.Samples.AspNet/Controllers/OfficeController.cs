using DotNetUtility.Samples.AspNet.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq.Expressions;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Mvc;

namespace DotNetUtility.Samples.AspNet.Controllers
{
    public class OfficeController : Controller
    {
        public ActionResult Index()
        {
            return View(CreateTestData1List());
        }
        public ActionResult ExportXlsx()
        {
            var dt = OfficeHelper.ToDatatable(CreateTestData1List(), new Dictionary<string, Expression<Func<TestData1, object>>> {
                { "Id",it=>it.Id },
                { "Name",it=>it.Name },
                { "Time",it=>it.Time}
            });
            var ds = new DataSet();
            ds.Tables.Add(dt);

            //直接输出到响应流以减少存取文件或MemoryStream的使用开销
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            ContentDispositionHeaderValue contentDisposition = new ContentDispositionHeaderValue("attachment");
            contentDisposition.FileName = "导出的Excel.xlsx";
            Response.Headers["Content-Disposition"] = contentDisposition.ToString();
            NPOIHelper.DatasetToExcel(ds, Response.OutputStream, ExcelFormat.Xlsx);
            return new EmptyResult();
        }
        [HttpPost]
        public ActionResult ImportXlsx(HttpPostedFile import)
        {
            DataSet ds = null;
            try
            {
                if (import != null && import.ContentLength > 0)
                {
                    ds = NPOIHelper.ExcelToDataset(import.InputStream);
                }
            }
            catch (Exception)
            {
                throw;
            }

            if (ds.Tables?.Count > 0)
            {
                var dt = ds.Tables[0];
                int successCount = 0;
                List<string> messageList = new List<string>(dt.Rows.Count);
                List<TestData1> testData1List = new List<TestData1>(dt.Rows.Count);
                foreach (var item in OfficeHelper.ConvertFromDatatable(dt, new Dictionary<string, Func<TestData1, object, DataRow, bool>> {
                    { "Id",(it,obj,row)=>OfficeHelper.TryConvertThen<int>(obj,v=>it.Id=v)},
                    { "Name",(it,obj,row)=>OfficeHelper.TryConvertThen<string>(obj,v=>it.Name=v)},
                    { "Time",(it,obj,row)=>OfficeHelper.TryConvertThen<DateTime>(obj,v=>it.Time=v)
                    }
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

                //下面部分只是为了展示出导入的数据
                ViewData["successCount"] = successCount;
                ViewData["messageList"] = messageList;
                ViewData["testData1List"] = testData1List;
                return View("Index", CreateTestData1List());
            }
            else
            {
                throw new Exception("所上传文件中不包含工作表！");
            }
        }
        List<TestData1> CreateTestData1List()
        {
            return new List<TestData1>()
            {
                new TestData1{
                    Id = 1,
                    Name = "张三",
                    Time = DateTime.Now.AddDays(-1)
                },
                new TestData1{
                    Id = 2,
                    Name = "李四",
                    Time = DateTime.Now.AddHours(-1)
                },
                new TestData1{
                    Id =3,
                    Name = "王五",
                    Time = DateTime.Now.AddMinutes(-1)
                }
            };
        }
    }
}