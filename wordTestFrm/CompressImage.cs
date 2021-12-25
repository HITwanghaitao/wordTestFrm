using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Drawing.Drawing2D;
using System.Linq.Expressions;

namespace wordTestFrm
{
    public class Resampler
    {
        /// <summary>
        /// Resamples all images in the document that are greater than the specified PPI (pixels per inch) to the specified PPI
        /// And converts them to JPEG with the specified quality setting.
        /// </summary>
        /// <param name="doc">The document to process.</param>
        /// <param name="desiredPpi">Desired pixels per inch. 220 high quality. 150 screen quality. 96 email quality.</param>
        /// <param name="jpegQuality">0 - 100% JPEG quality.</param>
        /// <returns></returns>
        public static int Resample(Document doc, int desiredPpi, int jpegQuality)
        {
            int count = 0;

            // Convert VML shapes.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                // It is important to use this method to correctly get the picture shape size in points even if the picture is inside a group shape.
                SizeF shapeSizeInPoints = shape.SizeInPoints;

                if (ResampleCore(shape.ImageData, shapeSizeInPoints, desiredPpi, jpegQuality))
                    count++;
            }

            return count;
        }

        /// <summary>
        /// Resamples one VML or DrawingML image
        /// </summary>
        public static bool ResampleCore(ImageData imageData, SizeF shapeSizeInPoints, int ppi, int jpegQuality)
        {
            // The are actually several shape types that can have an image (picture, ole object, ole control), let's skip other shapes.
            if (imageData == null)
                return false;

            // An image can be stored in the shape or linked from somewhere else. Let's skip images that do not store bytes in the shape.
            byte[] originalBytes = imageData.ImageBytes;
            if (originalBytes == null)
                return false;

            // Ignore metafiles, they are vector drawings and we don't want to resample them.
            ImageType imageType = imageData.ImageType;
            if (imageType.Equals(ImageType.Wmf) || imageType.Equals(ImageType.Emf))
                return false;

            try
            {
                double shapeWidthInches = ConvertUtil.PointToInch(shapeSizeInPoints.Width);
                double shapeHeightInches = ConvertUtil.PointToInch(shapeSizeInPoints.Height);

                // Calculate the current PPI of the image.
                ImageSize imageSize = imageData.ImageSize;
                double currentPpiX = imageSize.WidthPixels / shapeWidthInches;
                double currentPpiY = imageSize.HeightPixels / shapeHeightInches;

                Console.Write("Image PpiX:{0}, PpiY:{1}. ", (int)currentPpiX, (int)currentPpiY);

                // Let's resample only if the current PPI is higher than the requested PPI (e.g. we have extra data we can get rid of).
                if ((currentPpiX <= ppi) || (currentPpiY <= ppi))
                {
                    Console.WriteLine("Skipping.");
                    return false;
                }

                using (Image srcImage = imageData.ToImage())
                {
                    // Create a new image of such size that it will hold only the pixels required by the desired ppi.
                    int dstWidthPixels = (int)(shapeWidthInches * ppi);
                    int dstHeightPixels = (int)(shapeHeightInches * ppi);
                    using (Bitmap dstImage = new Bitmap(dstWidthPixels, dstHeightPixels))
                    {
                        // Drawing the source image to the new image scales it to the new size.
                        using (Graphics gr = Graphics.FromImage(dstImage))
                        {
                            gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            gr.DrawImage(srcImage, 0, 0, dstWidthPixels, dstHeightPixels);
                        }

                        // Create JPEG encoder parameters with the quality setting.
                        ImageCodecInfo encoderInfo = GetEncoderInfo(ImageFormat.Jpeg);
                        EncoderParameters encoderParams = new EncoderParameters();
                        encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, jpegQuality);

                        // Save the image as JPEG to a memory stream.
                        MemoryStream dstStream = new MemoryStream();
                        dstImage.Save(dstStream, encoderInfo, encoderParams);

                        // If the image saved as JPEG is smaller than the original, store it in the shape.
                        Console.WriteLine("Original size {0}, new size {1}.", originalBytes.Length, dstStream.Length);
                        if (dstStream.Length < originalBytes.Length)
                        {
                            dstStream.Position = 0;
                            imageData.SetImage(dstStream);
                            return true;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                // Catch an exception, log an error and continue if cannot process one of the images for whatever reason.
                Console.WriteLine("Error processing an image, ignoring. " + e.Message);
            }

            return false;
        }



        /// <summary>
        ///压缩图片并反馈Image
        /// </summary>
        public static Image ResampleCoreToImage(ImageData imageData, SizeF shapeSizeInPoints, int ppi, int jpegQuality)
        {
            // The are actually several shape types that can have an image (picture, ole object, ole control), let's skip other shapes.
            if (imageData == null)
                return null;

            // An image can be stored in the shape or linked from somewhere else. Let's skip images that do not store bytes in the shape.
            byte[] originalBytes = imageData.ImageBytes;
            if (originalBytes == null)
                return null;

            // Ignore metafiles, they are vector drawings and we don't want to resample them.
            ImageType imageType = imageData.ImageType;
            if (imageType.Equals(ImageType.Wmf) || imageType.Equals(ImageType.Emf))
                return null;

            try
            {
                double shapeWidthInches = ConvertUtil.PointToInch(shapeSizeInPoints.Width);
                double shapeHeightInches = ConvertUtil.PointToInch(shapeSizeInPoints.Height);

                // Calculate the current PPI of the image.
                ImageSize imageSize = imageData.ImageSize;
                double currentPpiX = imageSize.WidthPixels / shapeWidthInches;
                double currentPpiY = imageSize.HeightPixels / shapeHeightInches;

                Console.Write("Image PpiX:{0}, PpiY:{1}. ", (int)currentPpiX, (int)currentPpiY);

                // Let's resample only if the current PPI is higher than the requested PPI (e.g. we have extra data we can get rid of).
                if ((currentPpiX <= ppi) || (currentPpiY <= ppi))
                {
                    Console.WriteLine("Skipping.");
                    return null;
                }

                using (Image srcImage = imageData.ToImage())
                {
                    // Create a new image of such size that it will hold only the pixels required by the desired ppi.
                    int dstWidthPixels = (int)(shapeWidthInches * ppi);
                    int dstHeightPixels = (int)(shapeHeightInches * ppi);
                    using (Bitmap dstImage = new Bitmap(dstWidthPixels, dstHeightPixels))
                    {
                        // Drawing the source image to the new image scales it to the new size.
                        using (Graphics gr = Graphics.FromImage(dstImage))
                        {
                            gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            gr.DrawImage(srcImage, 0, 0, dstWidthPixels, dstHeightPixels);
                        }

                        // Create JPEG encoder parameters with the quality setting.
                        ImageCodecInfo encoderInfo = GetEncoderInfo(ImageFormat.Jpeg);
                        EncoderParameters encoderParams = new EncoderParameters();
                        encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, jpegQuality);

                        // Save the image as JPEG to a memory stream.
                        using (MemoryStream dstStream = new MemoryStream())
                        {
                            dstImage.Save(dstStream, encoderInfo, encoderParams);

                            // If the image saved as JPEG is smaller than the original, store it in the shape.
                            Console.WriteLine("Original size {0}, new size {1}.", originalBytes.Length, dstStream.Length);
                            if (dstStream.Length < originalBytes.Length)
                            {
                                dstStream.Position = 0;

                                return Image.FromStream(dstStream);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                // Catch an exception, log an error and continue if cannot process one of the images for whatever reason.
                Console.WriteLine("Error processing an image, ignoring. " + e.Message);
            }

            return null;
        }


        /// <summary>
        /// Gets the codec info for the specified image format. Throws if cannot find.
        /// </summary>
        private static ImageCodecInfo GetEncoderInfo(ImageFormat format)
        {
            ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();

            for (int i = 0; i < encoders.Length; i++)
            {
                if (encoders[i].FormatID == format.Guid)
                    return encoders[i];
            }

            throw new Exception("Cannot find a codec.");
        }


        /// <summary>
        /// 图片压缩
        /// </summary>
        /// <param name="flag">压缩比 1~100</param>
        /// <returns>内存流</returns>
        public static Stream GetPicThumbnail(Image iSource, int flag)
        {
            MemoryStream ms = null;
            ImageFormat tFormat = iSource.RawFormat;

            //以下代码为保存图片时，设置压缩质量  
            EncoderParameters ep = new EncoderParameters();
            long[] qy = new long[1];
            qy[0] = flag;//设置压缩的比例1-100  
            EncoderParameter eParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qy);
            ep.Param[0] = eParam;
            try
            {
                ImageCodecInfo[] arrayICI = ImageCodecInfo.GetImageEncoders();
                ImageCodecInfo jpegICIinfo = null;
                for (int x = 0; x < arrayICI.Length; x++)
                {
                    if (arrayICI[x].FormatDescription.Equals("JPEG"))
                    {
                        jpegICIinfo = arrayICI[x];
                        break;
                    }
                }
                if (jpegICIinfo != null)
                {
                    ms = new MemoryStream();
                    //iSource.Save(outPath, jpegICIinfo, ep);//dFile是压缩后的新路径  
                    iSource.Save(ms, jpegICIinfo, ep);
                }
                else
                {
                    ms = new MemoryStream();
                    iSource.Save(ms, tFormat);

                }
                return ms;
            }
            catch
            {
                return null;
            }
            finally
            {
                iSource.Dispose();
            }
        }


        /// <summary>
        /// 图片压缩
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ppi"></param>
        /// <param name="Quality"></param>
        public static int SetStyleForImage(Document doc, int ppi, int Quality,double width)
        {
            try
            {
                //获取所有图片
                NodeCollection nodes_Pic = doc.GetChildNodes(NodeType.Shape, true);
                int imageIndex = 0;
                for (int i = 0; i < nodes_Pic.Count; i++)
                {
                    Shape shape = (Shape)nodes_Pic[i];
                    if (shape.HasImage)
                    {
                        string time = DateTime.Now.ToString("HHmmssfff");
                        //扩展名
                        string ex = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                        //文件名
                        string imgName = string.Format("{0}_{1}{2}", time, imageIndex, ex);
                        Image img = shape.ImageData.ToImage();
                        //double iwidth = shape.Width;
                        //double height = shape.Height;
                          shape.HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment.Center;
                        Image newImage = Resampler.ResampleCoreToImage(shape.ImageData, shape.SizeInPoints, ppi, Quality);
                        if (newImage == null) continue;
                        shape.ImageData.SetImage(newImage);
                        double v = ConvertUtil.MillimeterToPoint(width * 10) / newImage.Width;
                        shape.Height = v * newImage.Height;
                        shape.Width = v * newImage.Width;
                        //shape.Width = width;
                        imageIndex++;
                    }
                }
                return 0;
            }
            catch (Aspose.Words.FileCorruptedException ex)
            {
                return -2;
            }
            catch (Exception ex)
            {
                return -1;
            }
          
        }

    }
}
