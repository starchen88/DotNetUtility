using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace DotNetUtility.Samples.AspNet.Controllers
{
    public class ImageController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public static string ThumbnailLocation = "~/Thumbnails/";
        public ActionResult Thumbnail(string url, int width, int height, string mode)
        {
            if (string.IsNullOrEmpty(mode))
            {
                mode = null;
            }
            else
            {
                switch (mode = mode.ToLower())
                {
                    case "cut":
                        break;
                    case "padding":
                        break;
                    default:
                        mode = null;
                        break;
                }
            }

            var rootPath = Server.MapPath("~/");
            var filepath = Server.MapPath(url);
            string thumbnailPath = Path.Combine(Server.MapPath(ThumbnailLocation), filepath.Replace(rootPath, string.Empty));
            thumbnailPath = thumbnailPath + "_" + width + "_" + height + "_" + mode + ".jpg";
            if (System.IO.File.Exists(thumbnailPath))
            {
                return File(thumbnailPath, "image/jpeg");
            }
            else
            {
                var thumbnailDir = Path.GetDirectoryName(thumbnailPath);
                if (!Directory.Exists(thumbnailDir))
                {
                    Directory.CreateDirectory(thumbnailDir);
                }
                using (var oldImage = new Bitmap(filepath))
                {
                    Bitmap newImage;
                    switch (mode)
                    {
                        case "cut":
                            newImage = ImageHelper.CreateThumbnailWithCut(oldImage, width, height);
                            break;
                        case "padding":
                            newImage = ImageHelper.CreateThumbnailWithPadding(oldImage, width, height);
                            break;
                        default:
                            newImage = ImageHelper.CreateThumbnail(oldImage, width, height);
                            break;
                    }
                    using (newImage)
                    {
                        int quality = 72;

                        //使用多线程来保存文件以提升性能表现
                        var saveTask = Task.Run(() =>
                          {
                              string tempFilename = thumbnailPath + Path.GetRandomFileName();
                              try
                              {
                                  //这里使用临时文件保存然后重命名而非直接写入是为了减少多线程争用导致的冲突
                                  ImageHelper.SaveAsJpeg(newImage, quality, tempFilename);
                                  System.IO.File.Move(tempFilename, thumbnailPath);
                              }
                              catch (Exception)
                              {
                                  throw;//TODO:此处需要处理异常
                              }
                              finally
                              {
                                  System.IO.File.Delete(tempFilename);//TODO:此处需要处理可能的异常
                              }
                          });
                        ImageHelper.SaveAsJpeg(newImage, quality, Response.OutputStream);

                        Task.WaitAll(saveTask);
                    }
                }
                return new EmptyResult();
            }
        }
    }
}