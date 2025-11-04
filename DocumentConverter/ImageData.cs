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

        public override string ToString()
        {
            return $"ImageData: {FileName} (RId: {RelationshipId}, Size: {Data?.Length ?? 0} bytes, ({Width}x{Height}))";
        }
    }
}