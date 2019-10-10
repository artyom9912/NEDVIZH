using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;
using HtmlAgilityPack;

namespace Project_parser
{
    public static class Farpost
    {
        /// <summary>
        /// Возвращает тип объявления
        /// </summary>
        /// <param name="link">Ссылка на объявление</param>
        /// <returns>Тип объявления (Поясн_Евгений.txt)</returns>
        public static string GetTypeByLink(string url)
        {
            if (url.Contains("realty/sell_business_realty"))
            {
                // Коммерческая недвижимость, продажа
                return "КН_пр";
            }
            if (url.Contains("realty/rent_business_realty"))
            {
                // Коммерческая недвижимость, аренда
                return "КН_ар";
            }
            else if (url.Contains("realty/land"))
            {
                // Земельные участки, продажа
                return "ЗУ_пр";
            }
            else if (url.Contains("realty/land-rent"))
            {
                // Земельные участки, аренда
                return "ЗУ_ар";
            }
            else if (url.Contains("realty/sale_garage"))
            {
                // Гараж, продажа
                return "ГАР_пр";
            }
            else if (url.Contains("realty/rent_garage"))
            {
                // Гараж, аренда
                return "ГАР_ар";
            }
            else if (url.Contains("realty/rent_flats") || url.Contains("realty/rent-apartment"))
            {
                // Квартира. Аренда и посуточно
                return "КВ_ар";
            }
            else if (url.Contains("realty/sell_flats"))
            {
                // Квартира, продажа
                return "КВ_пр";
            }
            else if (url.Contains("realty/rent_houses"))
            {
                // Дома и коттеджи, аренда
                return "ДМ_ар";
            }
            else if (url.Contains("realty/sell_houses"))
            {
                // Дома и коттеджи, продажа
                return "ДМ_пр";
            }
            else if (url.Contains("realty/dacha"))
            {
                // Дача, продажа
                return "ДЧ_пр";
            }
            return "??-??";
        }

        public static int GetItemsCount(string url)
        {
            int pages = -2;
            using (var wc = new WebClient())
            {
                wc.Headers.Add("User-Agent", General.DEFAULT_USER_AGENT);
                var byte_arr = wc.DownloadData(url);
                var page_cont = Encoding.Default.GetString(byte_arr);
                var idx = page_cont.IndexOf("предложения</strong>");
                if (idx == -1)
                {
                    idx = page_cont.IndexOf("предложений</strong>");
                    if (idx == -1)
                    {
                        idx = page_cont.IndexOf("предложение</strong>");
                        if (idx == -1)
                        {
                            pages = 0;
                        }
                    }
                }
                if (pages == -2)
                {
                    page_cont = page_cont.Remove(idx);
                    idx = page_cont.LastIndexOf("<strong>");
                    if (idx != -1)
                    {
                        page_cont = page_cont.Remove(0, idx + 8);
                        if (!Int32.TryParse(page_cont.Replace(" ", string.Empty).Trim(), out pages))
                        {
                            pages = 0;
                        }
                    }
                    else
                    {
                        pages = 0;
                    }
                }
            }
            return pages;
        }

        /// <summary>
        /// Возвращает список с инфой об объявах
        /// </summary>
        /// <param name="link">Ссылка на ленту</param>
        /// <returns>Список с инфой об объявах</returns>
        public static List<Advertisement> GetDataFromFeed(string url, string saveTo)
        {
            var result = new List<Advertisement>();
            HtmlDocument doc;
            var web = new HtmlWeb
            {
                AutoDetectEncoding = false,
                OverrideEncoding = Encoding.Default,
                UserAgent = General.DEFAULT_USER_AGENT
            };
            doc = web.Load(url);
            var cityDom = doc.DocumentNode.SelectNodes("//a[contains(@class, 'cityPop')]");
            var useAnnotations = url.Contains("primorskii-krai");
            var city = string.Empty;
            if(cityDom != null)
            {
                city = cityDom[0].InnerText.Trim();
            }
            else
            {
                city = string.Empty;
            }
            var advs = doc.DocumentNode.SelectNodes("//tr[contains(@class,'bull-item') and not(@data-accuracy)]");
            if(saveTo != null)
            {
                var content = doc.DocumentNode.InnerHtml;
                var linx = doc.DocumentNode.SelectNodes("//a[@href]");
                if(linx != null)
                {
                    foreach (var link in linx)
                    {
                        var val = link.Attributes["href"].Value.Trim();
                        if (!val.StartsWith("http://") &&
                           !val.StartsWith("https://"))
                        {
                            content = content.Replace('"' + val + '"', "\"https://www.farpost.ru" + val + '"');
                        }
                    }
                }
                File.WriteAllText(saveTo, content, Encoding.UTF8);
            }
            if (advs != null)
            {
                foreach (var adv in advs)
                {
                    try
                    {
                        var geo = string.Empty;
                        var title = WebUtility.HtmlDecode(adv.SelectSingleNode(".//td[@class='descriptionCell']/a[@name and @data-stat]").InnerText).Trim();
                        var geoDom = adv.SelectSingleNode(".//a[@data-geo]");
                        if(geoDom != null)
                        {
                            geo = geoDom.Attributes["data-geo"].Value.Trim();
                        }
                        var viewsDom = adv.SelectSingleNode(".//*[contains(@class,'views')]");
                        var views = 0;
                        if(viewsDom != null)
                        {
                            int.TryParse(viewsDom.InnerHtml, out views);
                        }
                        var number = adv.SelectSingleNode(".//a[@name!='']").Attributes["name"].Value.Trim();
                        var link = adv.SelectSingleNode(".//a[@href!='#']").Attributes["href"].Value.Trim();
                        // Цены иногда может и не быть...
                        var price = "n/a";
                        var priceDom = adv.SelectSingleNode(".//span[@data-role='price']");
                        if (priceDom != null)
                        {
                            price = WebUtility.HtmlDecode(priceDom.InnerText).Trim();
                            if (price.Contains("₽")) price = price.Replace("₽", "");
                        }
                        // Площадь. Также иногда может отсутствовать
                        string[] annotations = null;
                        var square = "n/a";
                        var squareDom = adv.SelectSingleNode(".//div[contains(@class, 'annotation')]");
                        if (squareDom != null)
                        {
                            annotations = squareDom.InnerText.Split(new string[] { ", "}, StringSplitOptions.RemoveEmptyEntries);
                            foreach(var sep in annotations)
                            {
                                if(sep.Contains("кв."))
                                {
                                    square = WebUtility.HtmlDecode(sep).Trim();
                                    break;
                                }
                            }
                        }
                        if(useAnnotations)
                        {
                            var altCity = string.Empty;
                            var cityDivDom = adv.SelectSingleNode(".//div[contains(@class, 'city')]");
                            if (cityDivDom != null)
                            {
                                altCity = cityDivDom.InnerText;
                            }
                            else
                            {
                                altCity = annotations[annotations.Length - 1];
                                var idx = altCity.IndexOf("м..");
                                if (idx != -1)
                                {
                                    altCity = altCity.Remove(0, idx + 3);
                                }
                            }
                            city = altCity.Trim();
                            General.WriteLog(city);
                        }
                        var adv_m = new Advertisement()
                        {
                            Number = number,
                            Link = "https://www.farpost.ru" + link,
                            Title = title,
                            Price = price,
                            Square = square,
                            Geo = geo,
                            Type = GetTypeByLink(link),
                            Views = views,
                            Annotations = annotations,
                            City = city,
                            CurLink = saveTo
                        };
                        result.Add(adv_m);
                    }
                    catch(Exception ex)
                    {
                        General.WriteLog(ex.ToString());
                    }
                }
            }
            /*
             * 
             */
            advs = doc.DocumentNode.SelectNodes("//td[contains(@class,'bull-item') and not(@data-accuracy)]");
            if(advs != null)
            {

                foreach (var adv in advs)
                {
                    try
                    {
                        var title = WebUtility.HtmlDecode(adv.SelectSingleNode(".//div[@class='title']/a").InnerText).Trim();
                        var geo = string.Empty;
                        var geoDom = adv.SelectSingleNode(".//a[@data-geo]");
                        if (geoDom != null)
                        {
                            geo = geoDom.Attributes["data-geo"].Value.Trim();
                        }
                        var viewsDom = adv.SelectSingleNode(".//*[contains(@class,'views')]");
                        var views = 0;
                        if (viewsDom != null)
                        {
                            int.TryParse(viewsDom.InnerHtml, out views);
                        }
                        var number = adv.SelectSingleNode(".//a[@name!='']").Attributes["name"].Value.Trim();
                        var link = adv.SelectSingleNode(".//a[@href!='#']").Attributes["href"].Value.Trim();
                        // Цены иногда может и не быть...
                        var price = "n/a";
                        var priceDom = adv.SelectSingleNode(".//span[@data-role='price']");
                        if (priceDom != null)
                        {
                            price = WebUtility.HtmlDecode(priceDom.InnerText).Trim();
                        }
                        // Площадь. Также иногда может отсутствовать
                        string[] annotations = null;
                        var square = "n/a";
                        var squareDom = adv.SelectSingleNode(".//div[contains(@class, 'annotation')]");
                        if (squareDom != null)
                        {
                            annotations = squareDom.InnerText.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var sep in annotations)
                            {
                                if (sep.Contains("кв."))
                                {
                                    square = WebUtility.HtmlDecode(sep).Trim();
                                    break;
                                }
                            }
                        }
                        var altCity = string.Empty;
                        if (useAnnotations)
                        {
                            var cityDivDom = adv.SelectSingleNode(".//div[contains(@class, 'city')]");
                            if (cityDivDom != null)
                            {
                                altCity = cityDivDom.InnerText;
                            }
                            else
                            {
                                altCity = annotations[annotations.Length - 1];
                                var idx = altCity.IndexOf("м..");
                                if (idx != -1)
                                {
                                    altCity = altCity.Remove(0, idx + 3);
                                }
                            }
                            city = altCity.Trim();
                        }
                        var adv_m = new Advertisement()
                        {
                            Number = number,
                            Link = "https://www.farpost.ru" + link,
                            Title = title,
                            Price = price,
                            Square = square,
                            Geo = geo,
                            Type = GetTypeByLink(link),
                            Views = views,
                            Annotations = annotations,
                            City = city
                        };
                        result.Add(adv_m);
                    }
                    catch(Exception ex)
                    {
                        General.WriteLog(ex.ToString());
                    }
                }
            }
            return result;
        }
    }
}
