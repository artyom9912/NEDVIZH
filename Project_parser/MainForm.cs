using System;
using System.Collections.Generic;
using System.Net;
using System.Threading;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using HtmlAgilityPack;

namespace Project_parser
{
    public partial class MainForm : Form
    {
        private DB _db = new DB(@"Data Source=NEWSQLPROD\S05_MSSQLSERV;Integrated Security=False;User ID=sql_prod_user;Password=8546!*jE&qs14;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
        private bool _stopFlag = false;
        private bool _archiveFlag = false;
        private string _lastSaveToPath = string.Empty;

        public MainForm()
        {
            InitializeComponent();
        }

        private void ParsePage(string url)
        {
            if (!_db.Connect())
            {
                MessageBox.Show(General.ERROR_DB_CONNECT);
                _db.Close();
                Invoke(new Action(() =>
                {
                    startButton.Text = "Начать";
                    startButton.Enabled = true;
                    statusLabel.Text = General.STATUS_DEFAULT;
                }));
                return;
            }
            /*
             * Трём '#' и всё после него, если таковой есть
             */
            var shard_idx = url.IndexOf("#");
            if (shard_idx > -1)
            {
                url = url.Remove(shard_idx);
            }
            //
            Invoke(new Action(() =>
            {
                textBox1.Text = url;
            }));
            //
            
            var lastId = _db.GetLastID() + 1;
            var siteType = General.GetSiteType(url);
            uint cnt = 0;
            if (siteType == SiteType.Farpost)
            {
                if (url.Contains("status=archive")) _archiveFlag = true;
                else _archiveFlag = false;
                var singleType = Farpost.GetTypeByLink(url);
                var dateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var folderName = dateTime + "_" + singleType + (_archiveFlag ? "_arc" : string.Empty);
                var workingDir = General.FARPOST_FOLDER + folderName + "\\";
                var Comment = (CommentBox.Text != String.Empty ? CommentBox.Text.Replace(" ", "_") : string.Empty);
                Directory.CreateDirectory(workingDir);
                List<Advertisement> advsData;
                var saveTo = workingDir + "L_" + dateTime + "_" + (_archiveFlag ? "_arc" : string.Empty) + singleType + ".html";
                var typeInt = _db.GetAdvTypeByString(singleType);
                if (url.Contains("?page=") || url.Contains("&page="))
                {
                    Invoke(new Action(() =>
                    {
                        statusLabel.Text = "Собираем информацию с ленты...";
                    }));
                    advsData = Farpost.GetDataFromFeed(url, saveTo);
                    _db.InsertPageInfo(saveTo, typeInt, lastId, lastId + advsData.Count - 1, DateTime.Now.ToString("yyyy.dd.MM hh:mm:ss"));
                }
                else
                {
                    var pages = Farpost.GetItemsCount(url);
                    pages = (int)Math.Ceiling((float)(pages / 50)) + 1;
                    if (pages != 1)
                    {
                        Invoke(new Action(() =>
                        {
                            statusLabel.Text = "Собираем информацию со всех страниц ленты (0/" + pages + ")...";
                        }));
                        advsData = new List<Advertisement>();
                        for (var i = 1; i <= pages; i++)
                        {
                            if (_stopFlag)
                            {
                                break;
                            }
                            saveTo = workingDir + "L_" + dateTime + "_" + (_archiveFlag ? "_arc" : string.Empty) + singleType + "_" + i + ".html";
                            var urlp = url + (url.Contains("?") ? "&" : "?") + "page=" + i;
                            var temp = Farpost.GetDataFromFeed(urlp, saveTo);
                            
                            advsData.AddRange(temp);
                            _db.InsertPageInfo(saveTo, typeInt, lastId, lastId + advsData.Count - 1, DateTime.Now.ToString("yyyy.dd.MM hh:mm:ss"));
                            Invoke(new Action(() =>
                            {
                                statusLabel.Text = "Собираем информацию со всех страниц ленты (" + i + "/" + pages + ")...";
                            }));
                            if (pages != i)
                            {
                                General.SmallDelay();
                            }
                        }
                    }
                    else
                    {
                        Invoke(new Action(() =>
                        {
                            statusLabel.Text = "Собираем информацию с ленты...";
                        }));
                        advsData = Farpost.GetDataFromFeed(url, saveTo);
                        _db.InsertPageInfo(saveTo, typeInt, lastId, lastId + advsData.Count - 1, DateTime.Now.ToString("yyyy.dd.MM hh:mm:ss"));
                    }
                }
                if (_stopFlag)
                {
                    _db.Close();
                    Invoke(new Action(() =>
                    {
                        startButton.Text = "Начать";
                        startButton.Enabled = true;
                        statusLabel.Text = General.STATUS_DEFAULT;
                    }));
                    return;
                }
                if (advsData == null || advsData.Count == 0)
                {
                    MessageBox.Show("Не удалось собрать объявления с указаной страницы.");
                    _db.Close();
                    Invoke(new Action(() =>
                    {
                        startButton.Text = "Начать";
                        startButton.Enabled = true;
                        statusLabel.Text = General.STATUS_DEFAULT;
                    }));
                    return;
                }
                cnt = 0;
                Invoke(new Action(() =>
                {
                    listBox1.Items.Clear();
                }));
                Invoke(new Action(() =>
                {
                    statusLabel.Text = "Обработка данных...";
                }));
                var todel = new List<Advertisement>();
                var updated = new List<int>();
                var trigger = checkBox1.Checked;
                int newId = lastId;
                for (var i = 0; i < advsData.Count; i++)
                {
                    if (_stopFlag)
                    {
                        break;
                    }                    
                    var id = _db.GetIdByNumber(advsData[i].Number);
                    if (id != -1) // Если объява с таким внутренним идом уже существует, то
                    {
                        // Получаем информацию о ней. И если цена и площадь остались прежними - удаляем из списка на инсерты данную великолепную запись.
                        var adv = _db.GetAdvPdata(id);
                        if ((!trigger && General.GetNumbers(adv.Price) == General.GetNumbers(advsData[i].Price) && General.GetNumbers(adv.Square) == General.GetNumbers(advsData[i].Square) ||
                            (trigger && General.GetNumbers(adv.Price) == General.GetNumbers(advsData[i].Price) && General.GetNumbers(adv.Square) == General.GetNumbers(advsData[i].Square) && _db.GetTownByMainId(id) == -1 && _db.GetTownByMainId(id) == 0)))
                        {
                            todel.Add(advsData[i]);
                            continue;
                        }
                        updated.Add(newId);
                    }
                    
                    _db.InsertRecord();
                    newId++;
                }
                foreach (var item in todel)
                {
                    advsData.Remove(item);
                }
                string distr = String.Empty;
                if (advsData.Count > 0)
                {
                    foreach (var advData in advsData)
                    {
                        if (_stopFlag)
                        {
                            break;
                        }
                        Invoke(new Action(() =>
                        {
                            listBox1.Items.Add(advData.Link);
                            countLabel.Text = "Кол-во объявлений: " + ++cnt;
                        }));
                        if (advData.Annotations != null)
                        {
                            foreach (string item in advData.Annotations)
                            {
                                if (item.Contains("р-н"))
                                {
                                    distr = item;
                                    break;
                                }
                            }
                        }
                        var dir = workingDir + lastId + "_" + advData.Number + "_" + DateTime.Now.ToString("yyyyddMM_hhmmss") + (_archiveFlag ? "_arc" : string.Empty);
                        var newdir = workingDir.Trim('\\') + "_" + advsData.Count + Comment;
                        int is_new = 1;
                        if (updated.Contains(lastId)) is_new = 0;
                        int source = 1;
                        if (_archiveFlag) source = 2;
                        _db.UpdateAdvData(lastId, is_new, source, advData.Number, advData.Link, "www.farpost.ru", newdir, advData.CurLink, advData.Title, advData.Price, advData.Square, advData.Geo, advData.City, distr, advData.Views, typeInt);
                        _db.InsertSecData(lastId);
                        Directory.CreateDirectory(dir);
                        lastId++;
                    }
                    if (singleType.Contains("КН_") && !(url.Contains("&flatType") || url.Contains("?flatType")))
                    {
                        var commercialTypes = new string[][] {
                            new [] { "outlet", "Торговая точка" },
                            new [] { "manufacture", "Производство" },
                            new [] { "office", "Офис" },
                            new [] { "storage", "Склад" },
                            new [] { "etc", "Другое" },
                        };
                        var advsDataAd = new List<Advertisement>();
                        foreach (var commercialType in commercialTypes)
                        {
                            var urlp = url + (url.Contains("?") ? "&" : "?") + "flatType%5B%5D=" + commercialType[0];
                            var pages = Farpost.GetItemsCount(urlp);
                            pages = (int)Math.Ceiling((float)(pages / 50)) + 1;
                            if (pages != 1)
                            {
                                for (var i = 1; i <= pages; i++)
                                {
                                    if (_stopFlag)
                                    {
                                        break;
                                    }
                                    var paged = urlp + (urlp.Contains("?") ? "&" : "?") + "page=" + i;
                                    var temp = Farpost.GetDataFromFeed(paged, null);
                                    advsDataAd.AddRange(temp);
                                    if (pages != i)
                                    {
                                        General.SmallDelay();
                                    }
                                }
                            }
                            else
                            {
                                advsDataAd = Farpost.GetDataFromFeed(urlp, null);
                            }
                            foreach (var advDataAd in advsDataAd)
                            {
                                _db.UpdateComType(advDataAd.Number, commercialType[1]);
                            }
                        }
                    }
                }
                var newPath = workingDir.Trim('\\') + "_" + advsData.Count + "_" + Comment;
                Directory.Move(workingDir, newPath);
                _lastSaveToPath = newPath;
            }
            else if (siteType == SiteType.Cian)
            {
                var links = Cian.GetLinksFromFeed(url);
                if (links == null)
                {
                    MessageBox.Show("Не удалось собрать ссылки на объявления с указаной страницы.");
                    _db.Close();
                    Invoke(new Action(() =>
                    {
                        startButton.Text = "Начать";
                        startButton.Enabled = true;
                        statusLabel.Text = General.STATUS_DEFAULT;
                    }));
                    return;
                }
                cnt = 0;
                Invoke(new Action(() =>
                {
                    listBox1.Items.Clear();
                }));
                foreach (var link in links)
                {
                    if (string.IsNullOrEmpty(link))
                    {
                        continue;
                    }
                    Invoke(new Action(() =>
                    {
                        listBox1.Items.Add(link);
                        countLabel.Text = "Кол-во объявлений: " + ++cnt;
                    }));
                    General.Delay();
                }
            }
            _db.Close();
            Invoke(new Action(() =>
            {
                startButton.Text = "Начать";
                startButton.Enabled = true;
                statusLabel.Text = General.STATUS_DEFAULT;
            }));
        }

        private void StartButton_Click(object sender, EventArgs e)
        {

            if (startButton.Text == "Начать")
            {
                //_archiveFlag = false;
                _stopFlag = false;
                countLabel.Text = "Кол-во объявлений: 0";
                startButton.Text = "Остановить";
                var processThread = new Thread(() => ParsePage(textBox1.Text))
                {
                    IsBackground = true
                };
                processThread.Start();
                //ParsePage(textBox1.Text);
                statusLabel.Text = "Запускаемся...";
            }
            else
            {
                _stopFlag = true;
                startButton.Enabled = false;
                statusLabel.Text = "Останавливаемся...";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            statusLabel.Text = General.STATUS_DEFAULT;
        }

        private void GuideButton_Click(object sender, EventArgs e)
        {
            Process.Start(@"E:\Parser_Test\Readme.docx");
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            var af = new ActionsForm(_db);
            af.textBox2.Text = _lastSaveToPath;
            af.ShowDialog();
        }

        private void ArchiveButton_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Contains("?status=archive") || textBox1.Text.Contains("&status=archive"))
            {
                textBox1.Text = textBox1.Text.Replace("?status=archive", "");
                textBox1.Text = textBox1.Text.Replace("&status=archive", "");
                _archiveFlag = false;
            }
            else
            {
                textBox1.Text += (textBox1.Text.Contains("?") ? "&" : "?") + "status=archive";
                _archiveFlag = true;
            }
        }
    }
}