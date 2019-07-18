using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;
namespace DotNetUtility.Tests
{
    [TestClass]
    public class ImageTest
    {
        string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "lena_std.tif");
        int[][] sizesFalse = new int[][] { new[] { -1, -2 }, new[] { 0, 10 } };
        int[][] sizes = new int[][] { new[] { -1, -2 }, new[] { 0, 10 }, new[] { 100, 80 }, new[] { 1024, 1024 }, /*new[] { int.MaxValue, int.MaxValue }*/ };
        [TestMethod]
        public void CreateThumbnailTest()
        {
            Console.WriteLine("CreateThumbnailTest");
            foreach (var size in sizes)
            {
                using (var memoryStream = new MemoryStream())
                {
                    using (var bitmap = new Bitmap(file))
                    {
                        Console.WriteLine(size[0] + "x" + size[1]);
                        try
                        {
                            using (var newImage = ImageHelper.CreateThumbnail(bitmap, size[0], size[1]))
                            {
                                ImageHelper.SaveAsJpeg(newImage, 72, memoryStream);
                            }
                        }
                        catch (Exception ex)
                        {
                            Assert.IsTrue(ex is ArgumentOutOfRangeException);
                            Console.WriteLine(ex.Message);
                        }
                    }
                    Console.WriteLine("生成文件尺寸：" + memoryStream.Length);
                }
            }
        }
        [TestMethod]
        public void CreateThumbnailWithPaddingTest()
        {
            Console.WriteLine("CreateThumbnailWithPaddingTest");
            foreach (var size in sizes)
            {
                using (var memoryStream = new MemoryStream())
                {
                    using (var bitmap = new Bitmap(file))
                    {
                        Console.WriteLine(size[0] + "x" + size[1]);
                        try
                        {
                            using (var newImage = ImageHelper.CreateThumbnailWithPadding(bitmap, size[0], size[1]))
                            {
                                ImageHelper.SaveAsJpeg(newImage, 72, memoryStream);
                            }
                        }
                        catch (Exception ex)
                        {
                            Assert.IsTrue(ex is ArgumentOutOfRangeException);
                            Console.WriteLine(ex.Message);
                        }
                    }
                    Console.WriteLine("生成文件尺寸：" + memoryStream.Length);
                }
            }
        }
        [TestMethod]
        public void CreateThumbnailWithCutTest()
        {
            Console.WriteLine("CreateThumbnailWithCutTest");
            foreach (var size in sizes)
            {
                using (var memoryStream = new MemoryStream())
                {
                    using (var bitmap = new Bitmap(file))
                    {
                        Console.WriteLine(size[0] + "x" + size[1]);
                        try
                        {
                            using (var newImage = ImageHelper.CreateThumbnailWithCut(bitmap, size[0], size[1]))
                            {
                                ImageHelper.SaveAsJpeg(newImage, 72, memoryStream);
                            }
                        }
                        catch (Exception ex)
                        {
                            Assert.IsTrue(ex is ArgumentOutOfRangeException);
                            Console.WriteLine(ex.Message);
                        }
                    }
                    Console.WriteLine("生成文件尺寸：" + memoryStream.Length);
                }
            }
        }
    }
}
