using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Logging;

// For more information on enabling MVC for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace DotNetUtility.Samples.AspNetCore.Controllers
{
    public class ImageController : Controller
    {
        private readonly IFileProvider _fileProvider;
        private readonly ILogger _logger;
        public ImageController(ILogger<ImageController> logger, IFileProvider fileProvider)
        {
            _logger = logger;
            _fileProvider = fileProvider;
        }
        public IActionResult Index()
        {
            return View();
        }
        public static string ThumbnailLocation = "Thumbnails";
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
            var rootPath = _fileProvider.GetFileInfo("/").PhysicalPath;
            var filepath = _fileProvider.GetFileInfo(url).PhysicalPath;
            string thumbnailPath = Path.Combine(rootPath, ThumbnailLocation, filepath.Replace(rootPath, string.Empty));
            thumbnailPath = thumbnailPath + "_" + width + "_" + height + "_" + mode + ".jpg";
            if (System.IO.File.Exists(thumbnailPath))
            {
                return PhysicalFile(thumbnailPath, "image/jpeg");
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
                            catch (Exception ex)
                            {
                                _logger.LogError(ex, "缩略图保存异常:" + tempFilename);
                            }
                            finally
                            {
                                try
                                {
                                    System.IO.File.Delete(tempFilename);
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogError(ex, "缩略图保存异常:" + tempFilename);
                                }
                            }
                        });
                        ImageHelper.SaveAsJpeg(newImage, quality, Response.Body);
                        Task.WaitAll(saveTask);
                    }
                }
                return new EmptyResult();
            }
        }
    }
}
