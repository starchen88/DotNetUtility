using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace DotNetUtility
{
    /// <summary>
    /// 基于GDI+的图像处理帮助类
    /// </summary>
    public static class ImageHelper
    {
        #region CreateThumbnail
        /// <summary>
        /// 创建非等比例缩放图
        /// </summary>
        /// <param name="oldImage">原始图</param>
        /// <param name="newWidth">宽度，必须大于0</param>
        /// <param name="newHeight">高度，必须大于0</param>
        /// <param name="interpolationMode">缩放算法</param>
        /// <returns>生成的缩放图</returns>
        public static Bitmap CreateThumbnail(this Bitmap oldImage, int newWidth, int newHeight, InterpolationMode interpolationMode = InterpolationMode.HighQualityBicubic)
        {
            if (newWidth <= 0 || newHeight <= 0)
            {
                throw new ArgumentOutOfRangeException("newWidth或newHeight不能小于0");
            }
            int oldWidth = oldImage.Width, oldHeight = oldImage.Height;
            Bitmap newImage = new Bitmap(newWidth, newHeight);
            using (Graphics g = Graphics.FromImage(newImage))
            {
                g.InterpolationMode = interpolationMode;
                g.DrawImage(oldImage, new Rectangle(0, 0, newWidth, newHeight));
            }
            return newImage;
        }
        /// <summary>
        /// 创建等比例缩放图，以比例更远一边为据填充另一边
        /// </summary>
        /// <param name="oldImage">原始图</param>
        /// <param name="newWidth">宽度，必须大于0</param>
        /// <param name="newHeight">高度，必须大于0</param>
        /// <param name="interpolationMode">缩放算法</param>
        /// <returns>生成的缩放图</returns>
        public static Bitmap CreateThumbnailWithPadding(this Bitmap oldImage, int newWidth, int newHeight, InterpolationMode interpolationMode = InterpolationMode.HighQualityBicubic)
        {
            return CreateThumbnailWithPadding(oldImage, newWidth, newHeight, Color.FromArgb(255, 255, 255), interpolationMode);
        }
        /// <summary>
        /// 创建等比例缩放图，以比例更远一边为据填充另一边
        /// </summary>
        /// <param name="oldImage">原始图</param>
        /// <param name="newWidth">宽度，必须大于0</param>
        /// <param name="newHeight">高度，必须大于0</param>
        /// <param name="paddingColor">填充颜色</param>
        /// <param name="interpolationMode">缩放算法</param>
        /// <returns>生成的缩放图</returns>
        public static Bitmap CreateThumbnailWithPadding(this Bitmap oldImage, int newWidth, int newHeight, Color paddingColor, InterpolationMode interpolationMode = InterpolationMode.HighQualityBicubic)
        {
            if (newWidth <= 0 || newHeight <= 0)
            {
                throw new ArgumentOutOfRangeException("newWidth或newHeight不能小于0");
            }
            int oldWidth = oldImage.Width, oldHeight = oldImage.Height;
            Bitmap newImage = new Bitmap(newWidth, newHeight);
            int w = newWidth, h = newHeight;
            if (oldWidth * newHeight > oldHeight * newWidth)
            {
                h = oldHeight * newWidth / oldWidth;
            }
            else if (oldWidth * newHeight < oldHeight * newWidth)
            {
                w = oldWidth * newHeight / oldHeight;
            }
            using (Graphics g = Graphics.FromImage(newImage))
            {
                g.InterpolationMode = interpolationMode;
                g.FillRectangle(new SolidBrush(paddingColor), new Rectangle(0, 0, newImage.Width, newImage.Height));
                g.DrawImage(oldImage, new Rectangle((newImage.Width - w) / 2, (newImage.Height - h) / 2, w, h),
                    new Rectangle(0, 0, oldImage.Width, oldImage.Height),
                    GraphicsUnit.Pixel);
            }
            return newImage;
        }
        /// <summary>
        /// 创建等比例缩放图，以比例更接近一边为据裁剪另一边
        /// </summary>
        /// <param name="oldImage">原始图</param>
        /// <param name="newWidth">宽度，必须大于0</param>
        /// <param name="newHeight">高度，必须大于0</param>
        /// <param name="interpolationMode">缩放算法</param>
        /// <returns>生成的缩放图</returns>
        public static Bitmap CreateThumbnailWithCut(this Bitmap oldImage, int newWidth, int newHeight, InterpolationMode interpolationMode = InterpolationMode.HighQualityBicubic)
        {
            if (newWidth <= 0 || newHeight <= 0)
            {
                throw new ArgumentOutOfRangeException("newWidth或newHeight不能小于0");
            }
            int oldWidth = oldImage.Width, oldHeight = oldImage.Height;
            Bitmap newImage = new Bitmap(newWidth, newHeight);
            if (oldWidth * newHeight > oldHeight * newWidth)
            {
                oldWidth = oldHeight * newWidth / newHeight;
            }
            else if (oldWidth * newHeight < oldHeight * newWidth)
            {
                oldHeight = oldWidth * newHeight / newWidth;
            }
            using (Graphics g = Graphics.FromImage(newImage))
            {
                g.InterpolationMode = interpolationMode;
                g.DrawImage(oldImage, new Rectangle(0, 0, newImage.Width, newImage.Height), new Rectangle((oldImage.Width - oldWidth) / 2, (oldImage.Height - oldHeight) / 2, oldWidth, oldHeight), GraphicsUnit.Pixel);
            }
            return newImage;
        }
        #endregion
        #region SaveAs        
        /// <summary>
        /// 保存Bitmap为jpeg到Stream
        /// </summary>
        /// <param name="bitmap">原始图</param>
        /// <param name="quality">压缩质量</param>
        /// <param name="stream">保存到的Stream</param>
        public static void SaveAsJpeg(this Bitmap bitmap, int quality, Stream stream)
        {
            ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);

            Encoder encoder = Encoder.Quality;
            EncoderParameters parms = new EncoderParameters(1);
            EncoderParameter parm = new EncoderParameter(encoder, (long)quality);
            parms.Param[0] = parm;
            bitmap.Save(stream, jgpEncoder, parms);
        }
        /// <summary>
        /// 保存Bitmap为jpeg到文件
        /// </summary>
        /// <param name="bitmap">原始图</param>
        /// <param name="quality">压缩质量</param>
        /// <param name="filename">保存到的文件路径</param>
        public static void SaveAsJpeg(this Bitmap bitmap, int quality, string filename)
        {
            ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);

            Encoder encoder = Encoder.Quality;
            EncoderParameters parms = new EncoderParameters(1);
            EncoderParameter parm = new EncoderParameter(encoder, (long)quality);
            parms.Param[0] = parm;
            bitmap.Save(filename, jgpEncoder, parms);
        }
        #endregion       
        public static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();

            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }
    }
}
