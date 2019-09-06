using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml;




namespace DocExtract
{
    class Program
    {
        static string headerStyle = @"    <head><style type='text/css'>" + 
            "div { position: absolute;height: 80%; width: 80%; top: 10%; left: 10%; }" +
            "pre {white-space: pre-wrap;  white-space: -moz-pre-wrap;  white-space: -pre-wrap; white-space: -o-pre-wrap; word-wrap: break-word;      }" +
            "</style></head>";

        static void Main(string[] args)
        {
            DirectoryInfo directory = new DirectoryInfo(@"C:\Documentation\");
            FileInfo[] allFiles = directory.GetFiles("*.docx", SearchOption.AllDirectories);

            //Remove hidden files
            var goodFiles = allFiles.Where(f => !f.Attributes.HasFlag(FileAttributes.Hidden)).OrderBy(f => f.Name);

            StringBuilder outputBody = new StringBuilder();

            //Build an Index
            outputBody.Append("<h1>Index</h1>");
            foreach (var file in goodFiles)
            {
                var name = Path.GetFileNameWithoutExtension(file.Name);
                outputBody.Append("<a href ='#" + name + "'>" + name + "</a><br/>");
            }

            foreach (var file in goodFiles)
            {
                string content = TextFromWord(file.FullName);

                //Cleanup content
                Regex.Replace(content, @"(\r\n){2,}", "\r\n\r\n");
                content = content.Replace(Environment.NewLine + Environment.NewLine, Environment.NewLine);

                //Encode for HTML
                content = HttpUtility.HtmlEncode(content).Replace("\n", "<br/>");

                //Add headers and metadata
                outputBody.Append("<a name = '" + Path.GetFileNameWithoutExtension(file.Name) + "'>");
                outputBody.Append("<h1>");
                outputBody.Append(Path.GetFileNameWithoutExtension(file.Name));
                outputBody.Append("</h1>");
                outputBody.Append("<p>Created: " + file.CreationTime + "</p>");
                outputBody.Append("<p>Edited: " + file.LastWriteTime + "</p>");
                outputBody.Append("<a href ='"+ file.FullName + "'>" + file.FullName + "</a>");

                //Add body
                outputBody.Append("<pre>" + content + "</pre>");


                outputBody.Append("<hr>");
            }

            var html = $"<html>{headerStyle}<body><div>" + outputBody.ToString() + "</div></body></html>";
            File.WriteAllText(@"C:\temp\doc.html",html);
            Process.Start(@"cmd.exe ", @"/c " + @"C:\temp\doc.html");
        }

        public static string TextFromWord(string file)
        {
            const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            StringBuilder textBuilder = new StringBuilder();
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(file, false))
            {
                // Manage namespaces to perform XPath queries.  
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("w", wordmlNamespace);

                // Get the document part from the package.  
                // Load the XML in the document part into an XmlDocument instance.  
                XmlDocument xdoc = new XmlDocument(nt);
                xdoc.Load(wdDoc.MainDocumentPart.GetStream());

                XmlNodeList paragraphNodes = xdoc.SelectNodes("//w:p", nsManager);
                foreach (XmlNode paragraphNode in paragraphNodes)
                {
                    XmlNodeList textNodes = paragraphNode.SelectNodes(".//w:t", nsManager);
                    foreach (System.Xml.XmlNode textNode in textNodes)
                    {
                        textBuilder.Append(textNode.InnerText);
                    }
                    textBuilder.Append(Environment.NewLine);
                }

            }
            return textBuilder.ToString();
        }
    }
}
