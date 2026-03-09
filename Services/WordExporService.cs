using HtmlToWordDemo.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToWordDemo.Services
{
    public class WordExportService
    {
        public void Export(DomTreeModel dom, CssMapperService cssMapper, string filePath)
        {
            using (var wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                wordDoc.AddMainDocumentPart();
                wordDoc.MainDocumentPart.Document = new Document(new Body());
                Body body = wordDoc.MainDocumentPart.Document.Body;

                foreach (var element in dom.Elements)
                {
                    var para = new Paragraph();
                    var style = cssMapper.MapElement(element);

                    Run run = new Run(new Text(element.InnerText));
                    run.RunProperties = style.RunProps;
                    para.Append(run);

                    body.Append(para);
                }

                wordDoc.MainDocumentPart.Document.Save();
            }
        }
    }
}