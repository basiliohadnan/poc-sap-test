using System.Drawing;
using System.Drawing.Imaging;
using ImageMagick;

namespace SAPTests.Helpers
{
    public class ImageEditor
    {
        public static Bitmap CropImage(Bitmap image, Rectangle cropRect)
        {
            var croppedImage = new Bitmap(cropRect.Width, cropRect.Height);
            using (var graphics = Graphics.FromImage(croppedImage))
            {
                graphics.DrawImage(image, new Rectangle(0, 0, cropRect.Width, cropRect.Height), cropRect, GraphicsUnit.Pixel);
            }
            return croppedImage;
        }

        public static Bitmap DuplicateSize(Image image)
        {
            var enlargedImage = new Bitmap(image.Width * 2, image.Height * 2);
            using (var graphics = Graphics.FromImage(enlargedImage))
            {
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                graphics.DrawImage(image, 0, 0, enlargedImage.Width, enlargedImage.Height);
            }
            return enlargedImage;
        }

        public static Bitmap ConvertToGrayscale(Bitmap image)
        {
            var grayscaleImage = new Bitmap(image.Width, image.Height, PixelFormat.Format24bppRgb);
            using (var graphics = Graphics.FromImage(grayscaleImage))
            {
                var colorMatrix = new ColorMatrix(new float[][]
                {
                    new float[] {0.299f, 0.299f, 0.299f, 0, 0},
                    new float[] {0.587f, 0.587f, 0.587f, 0, 0},
                    new float[] {0.114f, 0.114f, 0.114f, 0, 0},
                    new float[] {0, 0, 0, 1, 0},
                    new float[] {0, 0, 0, 0, 1}
                });
                using (var attributes = new ImageAttributes())
                {
                    attributes.SetColorMatrix(colorMatrix);
                    graphics.DrawImage(image, new Rectangle(0, 0, grayscaleImage.Width, grayscaleImage.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, attributes);
                }
            }
            return grayscaleImage;
        }

        public static Bitmap ApplyThreshold(Bitmap image, int threshold)
        {
            var thresholdedImage = new Bitmap(image.Width, image.Height);
            for (int x = 0; x < image.Width; x++)
            {
                for (int y = 0; y < image.Height; y++)
                {
                    Color originalColor = image.GetPixel(x, y);
                    Color thresholdedColor = Color.FromArgb(originalColor.R > threshold ? 255 : 0,
                                                             originalColor.G > threshold ? 255 : 0,
                                                             originalColor.B > threshold ? 255 : 0);
                    thresholdedImage.SetPixel(x, y, thresholdedColor);
                }
            }
            return thresholdedImage;
        }

        public static void SaveImage(Bitmap image, string imagePath)
        {
            string directory = Path.GetDirectoryName(imagePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(imagePath);
            string fileExtension = Path.GetExtension(imagePath);

            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string newFileName = $"{fileNameWithoutExtension}_{timestamp}{fileExtension}";

            string newImagePath = Path.Combine(directory, newFileName);

            image.Save(newImagePath, ImageFormat.Png);
        }

        public static string ConvertPngToJpg(string pngFilePath, string jpgFilePath)
        {
            using (MagickImage image = new MagickImage(pngFilePath))
            {
                // Convert PNG to JPG
                image.Write(jpgFilePath);
            }
            return jpgFilePath;
        }

        public static Bitmap InvertColors(Bitmap image)
        {
            Bitmap invertedImage = new Bitmap(image.Width, image.Height);

            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    Color originalColor = image.GetPixel(x, y);
                    Color invertedColor = Color.FromArgb(255 - originalColor.R, 255 - originalColor.G, 255 - originalColor.B);
                    invertedImage.SetPixel(x, y, invertedColor);
                }
            }

            return invertedImage;
        }
    }
}