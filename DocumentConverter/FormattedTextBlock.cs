namespace ClickUpDocumentImporter.DocumentConverter
{
    internal class FormattedTextBlock
    {
        public string Text { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float FontSize { get; set; }
        public string FontName { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderlined { get; set; }
        public bool IsStrikethrough { get; set; }
        public bool IsCode { get; set; }
        public bool IsHeading { get; set; }
        public bool IsBulletPoint { get; set; }
        public bool IsNumberedList { get; set; }
        public bool IsCodeBlock { get; set; }
        public bool IsBlockQuote { get; set; }
        public bool IsLink { get; set; }
        public string LinkUrl { get; set; }
        public string CodeLanguage { get; set; }
        public string Color { get; set; }

        public override string ToString()
        {
            return $"({X}, {Y}) [({FontSize}){FontName}] Text: {Text}";
        }
    }
}