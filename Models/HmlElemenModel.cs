namespace HtmlToWordDemo.Models
{
    public class HtmlElementModel
    {
        public string TagName { get; set; }
        public string InnerText { get; set; }
        public Dictionary<string, string> Attributes { get; set; } = new();
        public List<HtmlElementModel> Children { get; set; } = new();
    }

    public class DomTreeModel
    {
        public List<HtmlElementModel> Elements { get; set; } = new();
    }



}