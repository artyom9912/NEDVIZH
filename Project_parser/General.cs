using System;
using System.Collections.Generic;
using HtmlAgilityPack;
using System.Threading;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;


namespace Project_parser
{
    public struct Advertisement
    {
        public string Number;
        public string Link;
        public string Title;
        public string Price;
        public string Square;
        public string Geo;
        public string Type;
        public int Views;
        public string City;
        public string[] Annotations;
        public string CurLink;
    }

    public enum SiteType
    {
        Unknown,
        Farpost,
        Cian,
        Domclick
        //... и т.д.
    }

    public static class General
    {
        //public static string WORK_FOLDER = "E:\\Parser_Test\\";
        //public static string FARPOST_FOLDER = WORK_FOLDER + "Farpost\\";

        public static string WORK_FOLDER = "E:\\";
        public static string FARPOST_FOLDER = WORK_FOLDER + "00_obyav_Farpost_01_\\";

        public static string CIAN_FOLDER = WORK_FOLDER  + "Parser_Test\\" + "Cian\\";
        public static string LOG_PATH = WORK_FOLDER + "Parser_Test\\" + "log.txt";

        public const string STATUS_DEFAULT = "Нажмите кнопку \"Начать\"";
        public const string ERROR_DB_CONNECT = "Не удалось подключиться к БД";
        public const string DEFAULT_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36";

        
        public static void WriteLog(string text)
        {
            try
            {
                File.AppendAllText(LOG_PATH, DateTime.Now.ToString("[dd.MM.yy hh:mm:ss]") + " " + text + "\r\n");
            }
            catch
            {
            }
        }

        public static string GetNumbers(string input)
        {
            return new string(input.Where(c => char.IsDigit(c)).ToArray());
        }

        public static string ExtractValueFromDoc(SiteType type, string section, HtmlDocument doc)
        {
            var filename = WORK_FOLDER + "Parser_Test\\";
            switch (type)
            {
                default:
                    {
                        return string.Empty;
                    }
                case SiteType.Farpost:
                    {
                        filename += "farpost.ini";
                        break;
                    }
            }
            var manager = new IniReader(filename);
            var xpath = manager.GetValue("xpath", section, string.Empty).Trim('"');
            var idStr = manager.GetValue("id", section, string.Empty).Trim('"');
            if (string.IsNullOrEmpty(xpath))
            {
                return string.Empty;
            }
            if(doc == null)
            {
                return string.Empty;
            }
            var xp = doc.DocumentNode.SelectNodes(xpath);
            if(xp == null)
            {
                return string.Empty;
            }
            if (!int.TryParse(idStr, out int id))
            {
                if(idStr == "each")
                {
                    var seperator = manager.GetValue("seperator", section, string.Empty).Trim('"');
                    var result = string.Empty;
                    foreach (var x in xp)
                    {
                        result += x.InnerHtml.Trim() + seperator;
                    }
                    return StripHTML(result).Replace("&nbsp;", "").Trim();
                }
                else
                {
                    return string.Empty;
                }
            }
            if (xp.Count < id + 1)
            {
                return string.Empty;
            }
            return StripHTML(xp[id].InnerHtml).Replace("&nbsp;", "").Trim();
        }

        private static string StripHTML(string input)
        {
            return Regex.Replace(input, "<.*?>", String.Empty);
        }

        /// <summary>
        /// Ждун
        /// </summary>
        public static void Delay()
        {
            var rand = new Random();
            Thread.Sleep(rand.Next(35000, 45000));
        }

        /// <summary>
        /// Ждун
        /// </summary>
        public static void SmallDelay()
        {
            var rand = new Random();
            Thread.Sleep(rand.Next(2000, 7000));
        }

        /// <summary>
        /// Извлекает значение тега "title" из страницы
        /// </summary>
        /// <param name="content">HTML код страницы</param>
        /// <returns>Массив ссылок на объявы</returns>
        public static string ExtractTitleFromHtml(string content)
        {
            var idx = content.IndexOf("<title");
            if (idx == -1)
                return null;
            content = content.Remove(0, idx);
            idx = content.IndexOf(">");
            if (idx == -1)
                return null;
            content = content.Remove(0, idx + 1);
            idx = content.IndexOf("</title");
            if (idx == -1)
                return null;
            content = content.Remove(idx);
            return content;
        }

        /// <summary>
        /// Определяем тип сайта
        /// </summary>
        /// <param name="link">Ссылка на страницу</param>
        /// <returns>Возращает тип сайта</returns>
        public static SiteType GetSiteType(string link)
        {
            if (link.Contains("www.farpost."))
            {
                return SiteType.Farpost;
            }
            else if (link.Contains("cian.ru"))
            {
                return SiteType.Cian;
            }
            else if (link.Contains("www.farpost."))
            {
                return SiteType.Farpost;
            }
            return SiteType.Unknown;
        }

        /// <summary>
        /// Проверяем, валидный-ли URL
        /// </summary>
        /// <returns>Если удалось подключиться, то возвращает true. Иначе false</returns>
        public static bool IsValidUrl(string link)
        {
            try
            {
                new Uri(link);
                return true;
            }
            catch (UriFormatException)
            {
                return false;
            }
        }
    }
}
