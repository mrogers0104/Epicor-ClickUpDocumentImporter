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
        public string RelationshipId { get; set; }
        public byte[] Data { get; set; }
        public string FileName { get; set; }
        public int Index { get; set; }
        public string ContentType { get; set; }
        public int? PageNumber { get; set; } // For PDF images

        // position properties
        public float X { get; set; }
        public float Y { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }

    }
}
