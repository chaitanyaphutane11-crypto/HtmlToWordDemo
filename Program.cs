using System;
using HtmlToWordDemo.Services;

namespace HtmlToWordDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string htmlFile = "input.html";
            string wordFile = "output.docx";

            var parser = new HtmlParserService();
            var domTree = parser.Parse(htmlFile, out var cssFiles, out var embeddedCss);

            // Combine external + embedded CSS
            var allCssFiles = cssFiles.ToArray();
            var cssMapper = new CssMapperService("Config/cssRules.json", allCssFiles, embeddedCss);

            var exporter = new WordExportService();
            exporter.Export(domTree, cssMapper, wordFile);

            Console.WriteLine("✅ Conversion complete: " + wordFile);
        }
    }
}
