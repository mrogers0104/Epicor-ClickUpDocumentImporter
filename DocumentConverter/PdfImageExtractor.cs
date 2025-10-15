using ClickUpDocumentImporter.Helpers;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClickUpDocumentImporter.DocumentConverter
{

    /// <summary>
    /// Extract Images from PDF
    /// </summary>
    public class PdfImageExtractor
    {
        //public static List<ImageData> ExtractImagesFromPdf(string pdfFilePath)
        //{
        //    var images = new List<ImageData>();
        //    int imageIndex = 0;

        //    string uniqueId = Globals.CreateUniqueImageId(pdfFilePath);

        //    using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFilePath)))
        //    {
        //        int numberOfPages = pdfDoc.GetNumberOfPages();

        //        for (int pageNum = 1; pageNum <= numberOfPages; pageNum++)
        //        {
        //            var page = pdfDoc.GetPage(pageNum);
        //            var resources = page.GetResources();

        //            // !!! Need to find out why this line does not compile
        //            //var xObjects = resources.GetResourceNames(PdfName.XObject);
        //            //.Where(name => resources.GetResourceType(name) == PdfName.XObject);

        //            // Get the dictionary of XObjects associated with the page resources.
        //            var xObjectMap = resources.GetResource(PdfName.XObject);
        //            if (xObjectMap == null || !xObjectMap.IsDictionary())
        //            {
        //                // Skip if there are no XObjects on this page
        //                continue;
        //            }

        //            // Get the actual names (keys) from the XObject dictionary
        //            var xObjects = ((PdfDictionary)xObjectMap).KeySet();

        //            foreach (PdfName xObjectName in xObjects)
        //            {
        //                var xObject = resources.GetResource(xObjectName);

        //                if (xObject is PdfStream stream)
        //                {
        //                    var subType = stream.GetAsName(PdfName.Subtype);

        //                    if (PdfName.Image.Equals(subType))
        //                    {
        //                        try
        //                        {
        //                            PdfImageXObject image = new PdfImageXObject(stream);
        //                            byte[] imageBytes = image.GetImageBytes();

        //                            // Determine file extension
        //                            string extension = DetermineImageExtension(image);

        //                            images.Add(new ImageData
        //                            {
        //                                Data = imageBytes,
        //                                FileName = $"pdf_image_{imageIndex}_{uniqueId}_{extension}",
        //                                Index = imageIndex,
        //                                PageNumber = pageNum
        //                            });

        //                            imageIndex++;
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            Console.WriteLine($"Error extracting image on page {pageNum}: {ex.Message}");
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    Console.WriteLine($"Extracted {images.Count} images from PDF");
        //    return images;
        //}

        public static List<ImageData> ExtractImagesFromPdf(string pdfFilePath)
        {
            var images = new List<ImageData>();
            int imageIndex = 0;

            // Assuming this method exists and works
            string uniqueId = Globals.CreateUniqueImageId(pdfFilePath); 

            try
            {
                using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFilePath)))
                {
                    int numberOfPages = pdfDoc.GetNumberOfPages();

                    for (int pageNum = 1; pageNum <= numberOfPages; pageNum++)
                    {
                        var page = pdfDoc.GetPage(pageNum);
                        var resources = page.GetResources();

                        // 1. Get the XObject dictionary from the resources
                        PdfDictionary xObjectMap = resources.GetResource(PdfName.XObject);

                        // 2. Safely check if the dictionary exists
                        if (xObjectMap == null || !xObjectMap.IsDictionary())
                        {
                            continue; // Skip if no XObjects found
                        }

                        // 3. Iterate over the keys (names) in the XObject dictionary
                        var xObjects = ((PdfDictionary)xObjectMap).KeySet();

                        foreach (PdfName xObjectName in xObjects)
                        {
                            //// 3. Get the specific resource object using the correct resource type and name
                            //var xObject = resources.GetResource(PdfName.XObject, xObjectName);
                            // 3. Get the specific XObject (stream or dictionary) from the XObject dictionary
                            // This returns the PdfObject associated with the XObject name.
                            var xObject = xObjectMap.GetAsStream(xObjectName);

                            if (xObject != null && xObject.IsStream())
                            {
                                var stream = (PdfStream)xObject;
                                var subType = stream.GetAsName(PdfName.Subtype);

                                // 4. Check if the subtype is an Image
                                if (PdfName.Image.Equals(subType))
                                {
                                    try
                                    {
                                        // PdfImageXObject is the correct class to handle PDF image streams
                                        PdfImageXObject image = new PdfImageXObject(stream);
                                        byte[] imageBytes = image.GetImageBytes();

                                        // Determine file extension (assuming this helper method exists)
                                        string extension = DetermineImageExtension(image);

                                        images.Add(new ImageData
                                        {
                                            Data = imageBytes,
                                            FileName = $"pdf_image_{imageIndex}_{uniqueId}_{extension}",
                                            Index = imageIndex,
                                            PageNumber = pageNum
                                        });

                                        imageIndex++;
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error extracting image '{xObjectName}' on page {pageNum}: {ex.Message}");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred processing the PDF: {ex.Message}");
            }

            Console.WriteLine($"Extracted {images.Count} images from PDF");
            return images;
        }

        private static string DetermineImageExtension(PdfImageXObject image)
        {
            // Try to determine from filter
            var filter = image.GetPdfObject().GetAsName(PdfName.Filter);

            if (filter != null)
            {
                if (filter.Equals(PdfName.DCTDecode)) return ".jpg";
                if (filter.Equals(PdfName.JPXDecode)) return ".jp2";
                if (filter.Equals(PdfName.FlateDecode)) return ".png";
            }

            return ".png"; // Default
        }
    }
}
