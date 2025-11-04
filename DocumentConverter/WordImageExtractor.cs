using ClickUpDocumentImporter.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing; // For easy access to Drawing namespace elements
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing; // For Wordprocessing Drawing elements

namespace ClickUpDocumentImporter.DocumentConverter
{
    /// <summary>
    /// Extract Images from Word Document (.docx)
    /// </summary>
    public class WordImageExtractor
    {
        public static List<ImageData> ExtractImagesFromWord(string wordFilePath)
        {
            // Make sure these using directives are at the top of your file:
            // using DocumentFormat.OpenXml.Drawing;
            // using DocumentFormat.OpenXml.Wordprocessing;
            // using A = DocumentFormat.OpenXml.Drawing;
            // using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;
            // using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

            var images = new List<ImageData>();
            int imageIndex = 0;
            string uniqueId = Globals.CreateUniqueImageId(wordFilePath);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
            {
                var mainPart = wordDoc.MainDocumentPart;

                // --- Step 1: Find all Drawing elements in the document body ---
                var drawings = mainPart.Document.Body.Descendants<Drawing>();

                foreach (var drawing in drawings)
                {
                    // Initialize image data variables
                    string relationshipId = null;
                    long widthEmu = 0;
                    long heightEmu = 0;
                    long xOffsetEmu = 0; // xEmu
                    long yOffsetEmu = 0; // yEmu

                    // --- Case 1: Check for Floating Image (Anchor) ---
                    var anchorElement = drawing.Descendants<WP.Anchor>().FirstOrDefault();
                    if (anchorElement != null)
                    {
                        // 1. Get Relationship ID
                        var blip = anchorElement.Descendants<A.Blip>().FirstOrDefault();
                        if (blip?.Embed != null)
                        {
                            relationshipId = blip.Embed.Value;
                        }

                        // 2. Get Width and Height (Extent)
                        var extent = anchorElement.Descendants<A.Extents>().FirstOrDefault();
                        widthEmu = extent?.Cx ?? 0;
                        heightEmu = extent?.Cy ?? 0;

                        // 3. Get X and Y Position
                        // Horizontal Position (X)
                        var posH = anchorElement.Descendants<WP.HorizontalPosition>().FirstOrDefault();
                        if (posH != null && long.TryParse(posH.InnerText, out long resultH))
                        {
                            xOffsetEmu = resultH;
                        }
                        else
                        {
                            xOffsetEmu = 0;
                        }
                        //xOffsetEmu = posH?.Descendants<WP.PositionOffset>().FirstOrDefault()?.Text.ToLong() ?? 0;

                        // Vertical Position (Y)
                        var posV = anchorElement.Descendants<WP.VerticalPosition>().FirstOrDefault();
                        if (posV != null && long.TryParse(posV.InnerText, out long resultV))
                        {
                            yOffsetEmu = resultV;
                        }
                        else
                        {
                            yOffsetEmu = 0;
                        }

                        //yOffsetEmu = posV?.Descendants<WP.PositionOffset>().FirstOrDefault()?.Text.ToLong() ?? 0;
                    }
                    // --- Case 2: Check for Inline Image ---
                    else
                    {
                        var inlineElement = drawing.Descendants<WP.Inline>().FirstOrDefault();
                        if (inlineElement != null)
                        {
                            var blip = inlineElement.Descendants<A.Blip>().FirstOrDefault();
                            if (blip?.Embed != null)
                            {
                                relationshipId = blip.Embed.Value;
                            }

                            // Get Width and Height (Extent)
                            var extent = inlineElement.Descendants<A.Extents>().FirstOrDefault();
                            widthEmu = extent?.Cx ?? 0;
                            heightEmu = extent?.Cy ?? 0;

                            // X and Y remain 0 for inline images (relative position)
                            // If you need the x/y offset within the text run, you can find it,
                            // but the page position is usually zero unless complex layout is applied.
                        }
                    }

                    // --- Final Step: Extract Image Data (only if relationshipId was found) ---
                    if (relationshipId != null)
                    {
                        ImagePart imagePart = (ImagePart)mainPart.GetPartById(relationshipId);
                        images.Add(ExtractImageData(imagePart, imageIndex++, relationshipId, uniqueId, widthEmu, heightEmu, xOffsetEmu, yOffsetEmu));
                    }
                }
            }

            Console.WriteLine($"Extracted {images.Count} images from Word document");
            return images;
        }

        /// <summary>
        ///  Helper method to create the ImageData object
        /// </summary>
        /// <param name="imagePart"></param>
        /// <param name="index"></param>
        /// <param name="relationshipId"></param>
        /// <param name="uniqueId"></param>
        /// <param name="widthEmu"></param>
        /// <param name="heightEmu"></param>
        /// <param name="xEmu"></param>
        /// <param name="yEmu"></param>
        /// <returns></returns>
        private static ImageData ExtractImageData(ImagePart imagePart, int index, string relationshipId,
                                                  string uniqueId, long widthEmu, long heightEmu, long xEmu, long yEmu)
        {
            using (var stream = imagePart.GetStream())
            {
                using (var memoryStream = new MemoryStream())
                {
                    stream.CopyTo(memoryStream);
                    byte[] imageBytes = memoryStream.ToArray();

                    // GetImageExtension is a method you already have
                    string extension = GetImageExtension(imagePart.ContentType);

                    return new ImageData
                    {
                        RelationshipId = relationshipId,
                        Data = imageBytes,
                        FileName = $"image_{index}_{uniqueId}_{extension}",
                        Index = index,
                        ContentType = imagePart.ContentType,
                        Width = widthEmu.ConvertEMUtoInch(),
                        Height = heightEmu.ConvertEMUtoInch(),
                        X = xEmu.ConvertEMUtoInch(),
                        Y = yEmu.ConvertEMUtoInch()
                    };
                }
            }
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