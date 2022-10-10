namespace Word2DevOpsParser
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.Json;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    public static class WordParser
    {
        public static string Parse(Stream wordStream)
        {
            List<BacklogItem> backlogItemList = new List<BacklogItem>();

            using (WordprocessingDocument document = WordprocessingDocument.Open(wordStream, false))
            {
                Body body = document.MainDocumentPart.Document.Body;
                Dictionary<string, ParagraphStyle> styles = GetAvailableStyles(document.MainDocumentPart.StyleDefinitionsPart.Styles);

                BacklogItem backlogItem = new BacklogItem();

                foreach (Paragraph paragraph in body.Descendants<Paragraph>())
                {
                    string paragraphStyle = paragraph.ParagraphProperties?.ParagraphStyleId?.Val;
                    if (paragraphStyle != null)
                    {
                        ParagraphStyle primaryStyle = styles[paragraphStyle];

                        if (primaryStyle.Heading)
                        {
                            backlogItem = CreateNewBacklogItem(paragraph, primaryStyle);
                            backlogItemList.Add(backlogItem);
                            continue;
                        }
                    }

                    if (!ExtractPictures(document, backlogItem, paragraph))
                    {
                        backlogItem.Content += paragraph.InnerText + Environment.NewLine;
                    }
                }
            }

            return $"{{ \"items\": {JsonSerializer.Serialize(backlogItemList)}}}";
        }

        private static BacklogItem CreateNewBacklogItem(Paragraph paragraph, ParagraphStyle primaryStyle)
        {
            BacklogItem backlogItem = new BacklogItem();

            backlogItem.Name = paragraph.InnerText;
            backlogItem.StyleName = primaryStyle.Name;
            backlogItem.Indent = primaryStyle.Indent;
            backlogItem.Pictures = new Dictionary<string, string>();

            return backlogItem;
        }

        private static Dictionary<string, ParagraphStyle> GetAvailableStyles(Styles styles)
        {
            Dictionary<string, ParagraphStyle> styleDictionary = new Dictionary<string, ParagraphStyle>();

            foreach (Style style in styles.Descendants<Style>())
            {
                ParagraphStyle paragraphStyle = new ParagraphStyle();
                paragraphStyle.Name = style.StyleName.Val;
                paragraphStyle.Heading = style.StyleName.Val.ToString().StartsWith("heading", StringComparison.InvariantCultureIgnoreCase);
                paragraphStyle.Indent = style?.StyleParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val ?? 0;

                styleDictionary.Add(style.StyleId, paragraphStyle);
            }

            return styleDictionary;
        }

        private static bool ExtractPictures(WordprocessingDocument document, BacklogItem currParagraph, Paragraph paragraph)
        {
            bool pictureFound = false;

            foreach (Run run in paragraph.Descendants<Run>())
            {
                Drawing image = run.Descendants<Drawing>().FirstOrDefault();

                if (image != null)
                {
                    var imageFirst = image.Inline.Graphic.GraphicData.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().FirstOrDefault();
                    var blip = imageFirst.BlipFill.Blip.Embed.Value;
                    ImagePart img = (ImagePart)document.MainDocumentPart.Document.MainDocumentPart.GetPartById(blip);

                    string imgBase64 = ConvertToBase64(img.GetStream());

                    currParagraph.Content += blip + Environment.NewLine;
                    currParagraph.Pictures.Add(blip, imgBase64);
                    pictureFound = true;
                }
            }

            return pictureFound;
        }

        private static string ConvertToBase64(this Stream stream)
        {
            byte[] bytes;
            using (var memoryStream = new MemoryStream())
            {
                stream.CopyTo(memoryStream);
                bytes = memoryStream.ToArray();
            }

            return Convert.ToBase64String(bytes);
        }
    }
}
