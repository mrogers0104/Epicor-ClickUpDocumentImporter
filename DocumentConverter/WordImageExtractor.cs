using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClickUpDocumentImporter.DocumentConverter
{
    /// <summary>
    /// Extract Images from Word Document (.docx)
    /// </summary>
    public class WordImageExtractor
    {
        public static List<ImageData> ExtractImagesFromWord(string wordFilePath)
        {
            var images = new List<ImageData>();
            int imageIndex = 0;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
            {
                var mainPart = wordDoc.MainDocumentPart;

                // Get all image parts
                var imageParts = mainPart.ImageParts;

                foreach (var imagePart in imageParts)
                {
                    using (var stream = imagePart.GetStream())
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            stream.CopyTo(memoryStream);
                            byte[] imageBytes = memoryStream.ToArray();

                            // Get image extension from content type
                            string extension = GetImageExtension(imagePart.ContentType);

                            images.Add(new ImageData
                            {
                                Data = imageBytes,
                                FileName = $"image_{imageIndex}{extension}",
                                Index = imageIndex,
                                ContentType = imagePart.ContentType
                            });

                            imageIndex++;
                        }
                    }
                }
            }

            Console.WriteLine($"Extracted {images.Count} images from Word document");
            return images;
        }

        private static string GetImageExtension(string contentType)
        {
            return contentType switch
            {
                "image/png" => ".png",
                "image/jpeg" => ".jpg",
                "image/jpg" => ".jpg",
                "image/gif" => ".gif",
                "image/bmp" => ".bmp",
                "image/tiff" => ".tiff",
                _ => ".png"
            };
        }
    }
}
