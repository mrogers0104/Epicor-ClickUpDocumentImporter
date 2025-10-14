using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Xobject;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
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
        public static List<ImageData> ExtractImagesFromPdf(string pdfFilePath)
        {
            var images = new List<ImageData>();
            int imageIndex = 0;

            using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFilePath)))
            {
                int numberOfPages = pdfDoc.GetNumberOfPages();

                for (int pageNum = 1; pageNum <= numberOfPages; pageNum++)
                {
                    var page = pdfDoc.GetPage(pageNum);
                    var resources = page.GetResources();

                    // !!! Need to find out why this line does not compile
                    var xObjects = resources.GetResourceNames();
                        //.Where(name => resources.GetResourceType(name) == PdfName.XObject);

                    foreach (var xObjectName in xObjects)
                    {
                        var xObject = resources.GetResource(xObjectName);

                        if (xObject is PdfStream stream)
                        {
                            var subType = stream.GetAsName(PdfName.Subtype);

                            if (PdfName.Image.Equals(subType))
                            {
                                try
                                {
                                    PdfImageXObject image = new PdfImageXObject(stream);
                                    byte[] imageBytes = image.GetImageBytes();

                                    // Determine file extension
                                    string extension = DetermineImageExtension(image);

                                    images.Add(new ImageData
                                    {
                                        Data = imageBytes,
                                        FileName = $"pdf_image_{imageIndex}{extension}",
                                        Index = imageIndex,
                                        PageNumber = pageNum
                                    });

                                    imageIndex++;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error extracting image on page {pageNum}: {ex.Message}");
                                }
                            }
                        }
                    }
                }
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
