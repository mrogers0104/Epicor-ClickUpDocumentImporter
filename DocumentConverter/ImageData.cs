using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClickUpDocumentImporter.DocumentConverter
{
    // ===== Helper Class for Image Data =====
    public class ImageData
    {
        public byte[] Data { get; set; }
        public string FileName { get; set; }
        public int Index { get; set; }
        public string ContentType { get; set; }
        public int? PageNumber { get; set; } // For PDF images
    }
}
