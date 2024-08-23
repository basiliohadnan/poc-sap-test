using System.Drawing;
using System.Globalization;
using System.Text;
using Tesseract;

namespace SAPTests.Helpers
{
    public class OCRScanner
    {
        private static string tessPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "tessdata"));
        // Path to your template image containing all alphanumerics
        private static string templateStringImagePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "tessdata", "template-string.png"));
        public static string ExtractText(int roiX, int roiY, int roiWidth, int roiHeight, int threshold = 150, string imagePath = "", bool invertedColors = false)
        {
            if (imagePath == "")
            {
                imagePath = Global.processTest.CaptureWholeScreen();
            }

            // Load the template image
            Bitmap templateImage = new Bitmap(templateStringImagePath);

            using (var engine = new TesseractEngine(tessPath, "eng", EngineMode.Default))
            {
                using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    using (var image = Image.FromStream(stream))
                    {
                        var rect = new Rectangle(roiX, roiY, roiWidth, roiHeight);

                        // Convert the image to Bitmap
                        Bitmap bitmapImage = new Bitmap(image);

                        // Invert colors if requested
                        if (invertedColors)
                        {
                            bitmapImage = ImageEditor.InvertColors(bitmapImage);
                        }

                        // Convert the Bitmap image to grayscale
                        var grayscaleImage = ImageEditor.ConvertToGrayscale(bitmapImage);

                        // Apply thresholding
                        var thresholdedImage = ImageEditor.ApplyThreshold(grayscaleImage, threshold);

                        // Crop the thresholded image
                        using (var croppedImage = ImageEditor.CropImage(thresholdedImage, rect))
                        {
                            // Combine template image and cropped image horizontally
                            Bitmap combinedImage = CombineImages(templateImage, croppedImage);

                            // Save the combined image for inspection
                            string combinedImagePath = Path.Combine(Path.GetDirectoryName(imagePath), $"combined{threshold}_.png");
                            ImageEditor.SaveImage(combinedImage, combinedImagePath);

                            // Load the combined image directly into a Pix object
                            using (var memoryStream = new MemoryStream())
                            {
                                combinedImage.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                                memoryStream.Position = 0;
                                using (var pix = Pix.LoadFromMemory(memoryStream.ToArray()))
                                {
                                    // Perform OCR on the combined image
                                    using (var page = engine.Process(pix, PageSegMode.Auto))
                                    {
                                        var text = new StringBuilder();
                                        var targetRegionXStart = templateImage.Width;
                                        var targetRegionXEnd = targetRegionXStart + roiWidth;
                                        var iterator = page.GetIterator();
                                        iterator.Begin();

                                        do
                                        {
                                            if (iterator.TryGetBoundingBox(PageIteratorLevel.Word, out var boundingBox))
                                            {
                                                int x1 = boundingBox.X1;
                                                int y1 = boundingBox.Y1;
                                                int x2 = boundingBox.X2;
                                                int y2 = boundingBox.Y2;

                                                // Check if the bounding box of the word is within the target region
                                                if (x1 >= targetRegionXStart && x2 <= targetRegionXEnd)
                                                {
                                                    text.Append(iterator.GetText(PageIteratorLevel.Word));
                                                    text.Append(" ");
                                                }
                                            }
                                        } while (iterator.Next(PageIteratorLevel.Word));

                                        // Return the extracted text from the target region
                                        return text.ToString().Trim();
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private static Bitmap CombineImages(Bitmap templateImage, Bitmap targetImage)
        {
            // Create a new image wide enough to hold both the template and target images
            int combinedWidth = templateImage.Width + targetImage.Width;
            int combinedHeight = Math.Max(templateImage.Height, targetImage.Height);
            Bitmap combinedImage = new Bitmap(combinedWidth, combinedHeight);

            // Use graphics to draw both images on the combined image
            using (Graphics g = Graphics.FromImage(combinedImage))
            {
                g.DrawImage(templateImage, new Point(0, 0));
                g.DrawImage(targetImage, new Point(templateImage.Width, 0));
            }

            return combinedImage;
        }

        public static double TextToDouble(int x, int y, int width, int height, int threshold = 150, string imagePath = "", bool invertColors = false)
        {
            if (imagePath == "")
            {
                imagePath = Global.processTest.CaptureWholeScreen();
            }
            string ocr = ExtractText(x, y, width, height, threshold, imagePath, invertColors);
            ocr = StringHandler.ConvertCommasToDots(ocr);
            ocr = StringHandler.CleanInput(ocr);
            return double.Parse(ocr, CultureInfo.InvariantCulture);
        }
    }
}
