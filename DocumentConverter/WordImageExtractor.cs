using ClickUpDocumentImporter.Helpers;
using DocumentFormat.OpenXml.Packaging;
using HashidsNet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing; // For easy access to Drawing namespace elements
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing; // For Wordprocessing Drawing elements
using PIC = DocumentFormat.OpenXml.Drawing.Pictures; // For Picture elements

namespace ClickUpDocumentImporter.DocumentConverter
{
    /// <summary>
    /// Extract Images from Word Document (.docx)
    /// </summary>
    public class WordImageExtractor
    {
        //public static List<ImageData> ExtractImagesFromWord(string wordFilePath)
        //{
        //    var images = new List<ImageData>();
        //    int imageIndex = 0;

        //    string uniqueId = Globals.CreateUniqueImageId(wordFilePath);

        //    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
        //    {
        //        var mainPart = wordDoc.MainDocumentPart;

        //        RetrieveImageGeometry(mainPart);

        //        // Get all image parts
        //        var imageParts = mainPart.ImageParts;

        //        foreach (var imagePart in imageParts)
        //        {
        //            using (var stream = imagePart.GetStream())
        //            {
        //                using (var memoryStream = new MemoryStream())
        //                {
        //                    stream.CopyTo(memoryStream);
        //                    byte[] imageBytes = memoryStream.ToArray();

        //                    // Get image extension from content type
        //                    string extension = GetImageExtension(imagePart.ContentType);

        //                    ImageData imageData = new ImageData
        //                    {
        //                        Data = imageBytes,
        //                        FileName = $"image_{imageIndex}_{uniqueId}_{extension}",
        //                        Index = imageIndex,
        //                        ContentType = imagePart.ContentType
        //                    };

        //                    //RetrieveImageGeometry()

        //                    //images.Add(new ImageData
        //                    //{
        //                    //    Data = imageBytes,
        //                    //    FileName = $"image_{imageIndex}_{uniqueId}_{extension}",
        //                    //    Index = imageIndex,
        //                    //    ContentType = imagePart.ContentType
        //                    //});

        //                    images.Add(imageData);

        //                    imageIndex++;
        //                }
        //            }
        //        }
        //    }

        //    Console.WriteLine($"Extracted {images.Count} images from Word document");
        //    return images;
        //}

        //private static void RetrieveImageGeometry(MainDocumentPart mainPart)
        //{
        //    // Inside your WordprocessingDocument.Open block:
        //    var drawings = mainPart.Document.Body.Descendants<Drawing>();

        //    foreach (var drawing in drawings)
        //    {
        //        // 1. Find the Blip element which links to the imagePart (by RelationshipId)
        //        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();

        //        if (blip != null && blip.Embed != null)
        //        {
        //            string relationshipId = blip.Embed.Value;

        //            // 2. Resolve the relationshipId to the actual ImagePart
        //            ImagePart linkedImagePart = (ImagePart)mainPart.GetPartById(relationshipId);

        //            // Check if this is the image we are currently interested in (if iterating outside your current loop)
        //            // You would typically iterate over ALL drawings and then process the image data.

        //            // 3. Extract Width and Height from the Extent element
        //            var extent = drawing.Descendants<A.Extents>().FirstOrDefault();
        //            long widthEmu = 0;
        //            long heightEmu = 0;

        //            if (extent != null)
        //            {
        //                widthEmu = extent?.Cx ?? 0; // Width in EMUs
        //                heightEmu = extent?.Cy ?? 0; // Height in EMUs

        //                // To convert from EMUs to Pixels/Inches/CM, you need the DPI/DPC.
        //                // A common conversion for 96 DPI is Pixels = EMUs / 9525.
        //            }

        //            // 4. Extract X and Y for Floating Images (example using PositionOffset)
        //            // Position data is complex and depends heavily on whether the image is inline or floating.
        //            // For inline images (default), X and Y are often irrelevant or derived from layout.

        //            var positionOffset = drawing.Descendants<WP.PositionOffset>().FirstOrDefault();
        //            long xEmu = 0;
        //            long yEmu = 0;

        //            if (positionOffset != null)
        //            {
        //                // Floating image offset (X and Y coordinates)
        //                // Note: This often represents the offset from an anchor point, not absolute page coordinates.
        //                xEmu = positionOffset.Text;
        //            }

        //            // Now you have the image data AND its display dimensions/position!
        //            // You would now create your ImageData object with the retrieved values.
        //        }
        //    }
        //}

        //public static List<ImageData> ExtractImagesFromWord(string wordFilePath)
        //{
        //    // Make sure these using directives are at the top of your file:
        //    // using DocumentFormat.OpenXml.Drawing;
        //    // using DocumentFormat.OpenXml.Wordprocessing;
        //    // using A = DocumentFormat.OpenXml.Drawing;
        //    // using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;
        //    // using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

        //    var images = new List<ImageData>();
        //    int imageIndex = 0;
        //    string uniqueId = Globals.CreateUniqueImageId(wordFilePath);

        //    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
        //    {
        //        var mainPart = wordDoc.MainDocumentPart;

        //        // --- Step 1: Find all Drawing elements in the document body ---
        //        var drawings = mainPart.Document.Body.Descendants<Drawing>();

        //        foreach (var drawing in drawings)
        //        {
        //            // The image may be Inline or Floating (Anchor). We must check both structures.

        //            // Try to find the image part ID (Relationship ID)
        //            string relationshipId = null;

        //            // --- Case 1: Inline Image (e.g., image flows with text) ---
        //            var inlineElement = drawing.Descendants<WP.Inline>().FirstOrDefault();
        //            if (inlineElement != null)
        //            {
        //                var blip = inlineElement.Descendants<A.Blip>().FirstOrDefault();
        //                if (blip?.Embed != null)
        //                {
        //                    relationshipId = blip.Embed.Value;

        //                    // Extract Width/Height (Extent)
        //                    var extent = inlineElement.Descendants<A.Extents>().FirstOrDefault();
        //                    long widthEmu = extent?.Cx ?? 0;
        //                    long heightEmu = extent?.Cy ?? 0;

        //                    // X/Y are usually 0 for inline images
        //                    // You can optionally extract A.Offset from the Transform2D element if available, but they are often irrelevant for inline images
        //                    long xOffsetEmu = 0;
        //                    long yOffsetEmu = 0;

        //                    // --- Step 2: Extract Image Data and populate ImageData ---
        //                    if (relationshipId != null)
        //                    {
        //                        ImagePart imagePart = (ImagePart)mainPart.GetPartById(relationshipId);
        //                        images.Add(ExtractImageData(imagePart, imageIndex++, uniqueId,
        //                                                    widthEmu, heightEmu, xOffsetEmu, yOffsetEmu));
        //                    }
        //                }
        //            }

        //            // Note: Handling floating images (Anchor) is much more complex for X/Y position
        //            // as they are relative to the page margin, column, or paragraph.
        //            // For a basic extraction, focusing on Inline images (which cover most use cases) is often sufficient.
        //        }
        //    }

        //    Console.WriteLine($"Extracted {images.Count} images from Word document");
        //    return images;
        //}

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
                        } else
                        {
                            xOffsetEmu = 0;
                        }
                        //xOffsetEmu = posH?.Descendants<WP.PositionOffset>().FirstOrDefault()?.Text.ToLong() ?? 0;

                        // Vertical Position (Y)
                        var posV = anchorElement.Descendants<WP.VerticalPosition>().FirstOrDefault();
                        if (posV != null && long.TryParse(posV.InnerText, out long resultV))
                        {
                            yOffsetEmu = resultV;
                        } else
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
