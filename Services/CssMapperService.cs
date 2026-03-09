using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json;
using ExCSS;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToWordDemo.Models;

namespace HtmlToWordDemo.Services
{
    public class CssMapperService
    {
        private List<CssRule> _rules;
        private Dictionary<string, string> _classStyles = new();

        public CssMapperService(string configPath, string[] cssPaths, string embeddedCss)
        {
            // Load JSON config
            string json = File.ReadAllText(configPath);
            var config = JsonConvert.DeserializeObject<CssConfig>(json);
            _rules = config.Rules;

            var parser = new StylesheetParser();

            // Load external CSS files
            if (cssPaths != null)
            {
                foreach (var path in cssPaths)
                {
                    if (File.Exists(path))
                    {
                        var stylesheet = parser.Parse(File.ReadAllText(path));
                        foreach (var rule in stylesheet.StyleRules)
                        {
                            string selector = rule.SelectorText.Trim('.');
                            string declarations = rule.Style.ToString();
                            _classStyles[selector] = declarations; // later overrides earlier
                        }
                    }
                }
            }

            // Load embedded <style> CSS
            if (!string.IsNullOrEmpty(embeddedCss))
            {
                var stylesheet = parser.Parse(embeddedCss);
                foreach (var rule in stylesheet.StyleRules)
                {
                    string selector = rule.SelectorText.Trim('.');
                    string declarations = rule.Style.ToString();
                    _classStyles[selector] = declarations; // embedded overrides external
                }
            }
        }

        public WordStyleModel MapElement(HtmlElementModel element)
        {
            var props = new RunProperties();

            // Priority: Inline > Embedded/Class > External
            if (element.Attributes.ContainsKey("style"))
                ApplyRules(element.Attributes["style"], props);

            if (element.Attributes.ContainsKey("class"))
            {
                string className = element.Attributes["class"];
                if (_classStyles.ContainsKey(className))
                    ApplyRules(_classStyles[className], props);
            }

            return new WordStyleModel { RunProps = props };
        }

        private void ApplyRules(string styleText, RunProperties props)
        {
            foreach (var rule in _rules)
            {
                if (styleText.Contains(rule.Selector))
                {
                    switch (rule.WordStyle.Type)
                    {
                        case "Color":
                            props.Append(new Color() { Val = rule.WordStyle.Value });
                            break;
                        case "Bold":
                            props.Append(new Bold());
                            break;
                        case "Italic":
                            props.Append(new Italic());
                            break;
                        case "Shading":
                            props.Append(new Shading()
                            {
                                Val = ShadingPatternValues.Clear,
                                Color = "auto",
                                Fill = rule.WordStyle.Fill
                            });
                            break;
                        case "FontSize":
                            props.Append(new FontSize() { Val = rule.WordStyle.Value });
                            break;
                        case "Alignment":
                            props.Append(new Justification()
                            {
                                Val = (JustificationValues)Enum.Parse(typeof(JustificationValues), rule.WordStyle.Value, true)
                            });
                            break;
                    }
                }
            }
        }
    }

    // Config models
    public class CssConfig
    {
        public List<CssRule> Rules { get; set; }
    }

    public class CssRule
    {
        public string Selector { get; set; }
        public WordStyle WordStyle { get; set; }
    }

    public class WordStyle
    {
        public string Type { get; set; }
        public string Value { get; set; }
        public string Fill { get; set; }
    }
}