using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OnenoteAddin
{
    public class OneNoteOperator
    {
        private Application oneNoteApp;

        public OneNoteOperator(Application app)
        {
            this.oneNoteApp = app;
        }

        public string GetSelectedText()
        {
            var sb = new StringBuilder();
            oneNoteApp.GetPageContent(oneNoteApp.Windows.CurrentWindow.CurrentPageId,
                                       out string xml,
                                       PageInfo.piSelection,
                                       XMLSchema.xs2013);

            var doc = new XmlDocument();
            doc.LoadXml(xml);

            var textNodes = doc.GetElementsByTagName("one:T");
            foreach (XmlNode node in textNodes)
            {
                if (node.Attributes["selected"] != null)
                {
                    var text = Regex.Replace(
                        node.InnerText,
                        "<.*?>",
                        string.Empty,
                        RegexOptions.Singleline);
                    sb.Append(text + "\n");
                }
            }

            return sb.ToString();
        }

        public void ReplaceSelectedText(string newText)
        {
            string pageId = oneNoteApp.Windows.CurrentWindow.CurrentPageId;
            oneNoteApp.GetPageContent(pageId, out string xml, PageInfo.piSelection, XMLSchema.xs2013);

            var doc = new XmlDocument();
            doc.LoadXml(xml);

            var textNodes = doc.GetElementsByTagName("one:T");
            var selectedNodes = new List<XmlNode>();
            foreach (XmlNode node in textNodes)
            {
                if (node.Attributes["selected"] != null)
                {
                    selectedNodes.Add(node);
                }
            }

            if (selectedNodes.Count > 0)
            {
                selectedNodes[0].InnerText = newText;
                for (int i = 1; i < selectedNodes.Count; i++)
                {
                    selectedNodes[i].InnerText = string.Empty;
                }
            }

            oneNoteApp.UpdatePageContent(doc.OuterXml, DateTime.MinValue, XMLSchema.xs2013);
        }

        public void ReplaceSelectedTextWithHtmlBlock(string style, string body)
        {
            string pageId = oneNoteApp.Windows.CurrentWindow.CurrentPageId;
            oneNoteApp.GetPageContent(pageId, out string xml, PageInfo.piSelection, XMLSchema.xs2013);

            var doc = new XmlDocument();
            doc.LoadXml(xml);

            var textNodes = doc.GetElementsByTagName("one:T");
            var selectedNodes = new List<XmlNode>();
            foreach (XmlNode node in textNodes)
            {
                if (node.Attributes["selected"] != null)
                {
                    selectedNodes.Add(node);
                }
            }

            if (selectedNodes.Count > 0)
            {
                // one:HTMLBlockノードを作成
                var htmlBlockNode = doc.CreateElement("one:HTMLBlock", doc.DocumentElement.NamespaceURI);
                var dataNode = doc.CreateElement("one:Data", doc.DocumentElement.NamespaceURI);
                dataNode.InnerXml = $"<![CDATA[<html><head><style>{style}</style></head><body>{body}</body></html>]]>";
                htmlBlockNode.AppendChild(dataNode);

                // one:Tノードの親のone:OEノードを置き換え
                var oeNode = selectedNodes[0].ParentNode;
                oeNode.ParentNode.ReplaceChild(htmlBlockNode, oeNode);

                for (int i = 1; i < selectedNodes.Count; i++)
                {
                    var node = selectedNodes[i].ParentNode;
                    node.ParentNode.RemoveChild(node);
                }
            }

            oneNoteApp.UpdatePageContent(doc.OuterXml, DateTime.MinValue, XMLSchema.xs2013);
        }
    }
}
