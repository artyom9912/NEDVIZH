using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using HtmlAgilityPack;

namespace Project_parser
{
    public static class Cian
    {
        public static string[] GetLinksFromFeed(string link)
        {
            var result = new List<string>();
            HtmlAgilityPack.HtmlDocument doc;
            var web = new HtmlWeb
            {
                AutoDetectEncoding = false,
                OverrideEncoding = Encoding.UTF8,
                UserAgent = General.DEFAULT_USER_AGENT
            };
            doc = web.Load(link);
            var nodes = doc.DocumentNode.SelectNodes("//div[@class='c6e8ba5398--main--1lHg7']//div[not(@*)]/a[@href!='#']");
            if(nodes != null)
            {
                foreach(var node in nodes)
                {
                    result.Add(node.Attributes["href"].Value);
                }
            }
            File.WriteAllText(Path.Combine(General.CIAN_FOLDER, "Feed", "l_" + DateTime.Now.ToString("ddMMyy_HHmmss") + "_cian.html"), doc.DocumentNode.InnerHtml);
            return result.ToArray();
        }
    }
}