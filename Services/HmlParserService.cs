using HtmlAgilityPack;
using System.Collections.Generic;
using System.Linq;
using HtmlToWordDemo.Models;

namespace HtmlToWordDemo.Services
{
    public class HtmlParserService
    {
        public DomTreeModel Parse(string filePath, out List<string> cssFiles, out string embeddedCss)
        {
            var doc = new HtmlDocument();
            doc.Load(filePath);

            // Collect external CSS files
            cssFiles = doc.DocumentNode
                .SelectNodes("//link[@rel='stylesheet']")
                ?.Select(n => n.GetAttributeValue("href", ""))
                .Where(href => !string.IsNullOrEmpty(href))
                .ToList() ?? new List<string>();

            // Collect embedded <style> blocks
            embeddedCss = string.Join("\n",
                doc.DocumentNode.SelectNodes("//style")?.Select(n => n.InnerText) ?? new List<string>());

            var domTree = new DomTreeModel();
            foreach (var node in doc.DocumentNode.SelectNodes("//body/*"))
            {
                domTree.Elements.Add(ParseNode(node));
            }
            return domTree;
        }

        private HtmlElementModel ParseNode(HtmlNode node)
        {
            var element = new HtmlElementModel
            {
                TagName = node.Name,
                InnerText = node.InnerText,
                Attributes = node.Attributes.ToDictionary(a => a.Name, a => a.Value)
            };

            foreach (var child in node.ChildNodes)
            {
                if (child.NodeType == HtmlNodeType.Element)
                    element.Children.Add(ParseNode(child));
            }

            return element;
        }
    }
}