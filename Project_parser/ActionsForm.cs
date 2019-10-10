using System;
using System.Net;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Globalization;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Project_parser
{
    public partial class ActionsForm : Form
    {
        private int _idMin = 0, _idMax = 0;
        private DB _db;
        private const string STATUS_BAR_DEFAULT = "Ожидание";
        private bool _stopFlag = false;

        public ActionsForm(DB db)
        {
            InitializeComponent();
            _db = db;
        }

        private void UploadAdvs()
        {
            if (!_db.Connect())
            {
                MessageBox.Show(General.ERROR_DB_CONNECT);
                _db.Close();
                Invoke(new Action(() =>
                {
                    statusLabel.Text = STATUS_BAR_DEFAULT;
                    uploadButton.Enabled = true;
                    exploreButton.Enabled = true;
                    parseButton.Enabled = true;
                    exportButton.Enabled = true;
                    textBox2.Enabled = true;
                    stopButton.Enabled = false;
                    fromBox.Enabled = true;
                    toBox.Enabled = true;
                }));
                return;
            }

            var targets = textBox2.Text.Split(';');
            var i = 0;
            foreach(var target in targets)
            {
                var paths = Directory.GetDirectories(target);
                Invoke(new Action(() =>
                {
                    statusLabel.Text = "Инициализация Selenium...";
                }));
                IWebDriver browser;
                var option = new ChromeOptions();
                option.AddArgument("--headless");
                option.AddArgument("--user-agent=" + General.DEFAULT_USER_AGENT);
                //option.AddArgument("--proxy-server=socks5://109.234.35.41:8888");
                var service = ChromeDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true;
                browser = new ChromeDriver(service, option);
                // Нужно сделать проверку является-ли фарпостовской сессией
                var defaultPageAddress = "https://www.farpost.ru/vladivostok/realty/sell_flats/?page=1";
                browser.Navigate().GoToUrl(defaultPageAddress);
                var def = (int)(toBox.Value - fromBox.Value) + 1;
                foreach (var path in paths)
                {
                    if (_stopFlag)
                    {
                        break;
                    }
                    var pathSplit = path.Split('\\');
                    var advName = pathSplit[pathSplit.Length - 1];
                    var nlink = target + "\\" + advName;
                    if (File.Exists(nlink + ".html") || File.Exists(nlink + "_arch.html"))
                    {
                        i++;
                        continue;
                    }
                    var idStr = pathSplit[pathSplit.Length - 1].Split('_')[0];
                    if (!int.TryParse(idStr, out int id))
                    {
                        i++;
                        continue;
                    }
                    if (id < fromBox.Value || id > toBox.Value)
                    {
                        i++;
                        continue;
                    }
                    Invoke(new Action(() =>
                    {
                        statusLabel.Text = "Загрузка объявления #" + idStr + " [" + i + "/" + def + "]";
                    }));
                    var link = _db.GetLinkById(id);
                    while (true)
                    {
                        if (_stopFlag)
                        {
                            break;
                        }
                        try
                        {
                            browser.Navigate().GoToUrl(link);
                            if (browser.PageSource.Contains("Из вашей подсети наблюдается подозрительная активность"))
                            {
                                _db.Close();
                                Invoke(new Action(() =>
                                {
                                    statusLabel.Text = STATUS_BAR_DEFAULT;
                                    uploadButton.Enabled = true;
                                    exploreButton.Enabled = true;
                                    parseButton.Enabled = true;
                                    exportButton.Enabled = true;
                                    textBox2.Enabled = true;
                                    stopButton.Enabled = false;
                                    fromBox.Enabled = true;
                                    toBox.Enabled = true;
                                }));
                                service.Dispose();
                                browser.Dispose();
                                MessageBox.Show("Капча! Завершаем работу.");
                                return;
                            }
                            break;
                        }
                        catch (Exception ex)
                        {
                            General.WriteLog(ex.Message);
                            Thread.Sleep(10000);
                        }
                    }
                    if (!browser.PageSource.Contains("Объявление находится в архиве и может быть неактуальным."))
                    {
                        for (var k = 0; k < 5; k++)
                        {
                            if (_stopFlag)
                            {
                                break;
                            }
                            try
                            {
                                var element = browser.FindElement(By.PartialLinkText("Показать контакты"));
                                element.Click();
                                var isf = false;
                                for (var ti = 0; ti < 20 && !_stopFlag; ti++)
                                {
                                    try
                                    {
                                        browser.FindElement(By.ClassName("dummy-listener_new-contacts"));
                                        isf = true;
                                        break;
                                    }
                                    catch
                                    {
                                        Thread.Sleep(500);
                                    }
                                }
                                if (isf)
                                {
                                    break;
                                }
                            }
                            catch (Exception ex)
                            {
                                General.WriteLog(ex.Message);
                                Thread.Sleep(1000);
                            }
                        }
                    }
                    else
                    {
                        nlink += "_arch";
                    }
                    try
                    {
                        var element = browser.FindElement(By.ClassName("expand--button"));
                        element.Click();
                        for (var ti = 0; ti < 20 && !_stopFlag; ti++)
                        {
                            try
                            {
                                browser.FindElement(By.ClassName("mod__active"));
                                break;
                            }
                            catch
                            {

                            }
                            Thread.Sleep(1000);
                        }
                    }
                    catch (Exception)
                    {
                    }
                    var content = browser.PageSource;


                    var images = new List<string>();
                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(content);
                    var nodes = doc.DocumentNode.SelectNodes("//img");
                    if (nodes != null)
                    {
                        var nodesArray = nodes.ToArray();
                        foreach (var node in nodesArray)
                        {
                            images.Add(node.Attributes["src"].Value);
                        }
                    }

                    int img_name = 0;
                    using (var wc = new WebClient())
                    {
                        foreach (var imageSrc in images)
                        {
                            if (_stopFlag)
                            {
                                break;
                            }
                            try
                            {
                                var ext = imageSrc;
                                var idx = ext.LastIndexOf("/");
                                if (idx == -1)
                                {
                                    continue;
                                }
                                ext = ext.Remove(0, idx + 1);
                                idx = ext.LastIndexOf(".");
                                ext = idx != -1 ? ext.Remove(0, idx) : ".jpg";
                                var pictureName = img_name.ToString() + ext;

                                content = content.Replace(imageSrc, Path.Combine(advName, pictureName));
                                wc.DownloadFile(imageSrc, Path.Combine(path, pictureName));
                                img_name++;
                            }
                            catch (Exception ex)
                            {
                                General.WriteLog(ex.Message);
                            }
                        }
                    }
                    nodes = doc.DocumentNode.SelectNodes("//*[href]");
                    if (nodes != null)
                    {
                        var nodesArray = nodes.ToArray();
                        foreach (var node in nodesArray)
                        {
                            var value = node.Attributes["href"].Value;
                            if (!value.StartsWith("http://") && !value.StartsWith("https://"))
                            {
                                var newValue = value;
                                if(newValue.StartsWith("/"))
                                {
                                    newValue.Remove(0, 1);
                                }
                                content = content.Replace(value, "https://www.farpost.ru/" + newValue);
                            }
                        }
                    }

                    var htmled = nlink + ".html";
                    File.WriteAllText(htmled, content, Encoding.UTF8);
                    FileInfo fileInfo = new FileInfo(htmled);
                    string newname;
                    if(fileInfo.Length <= 25600)
                    {
                        newname = nlink + "_udalen.html";
                        File.Move(htmled, newname);
                    }
                    _db.ChangeAdvNLink(id, htmled);
                    General.Delay();
                    i++;
                }
                _db.Close();
                Invoke(new Action(() =>
                {
                    statusLabel.Text = STATUS_BAR_DEFAULT;
                    uploadButton.Enabled = true;
                    exploreButton.Enabled = true;
                    parseButton.Enabled = true;
                    exportButton.Enabled = true;
                    textBox2.Enabled = true;
                    stopButton.Enabled = false;
                    fromBox.Enabled = true;
                    toBox.Enabled = true;
                }));
                service.Dispose();
                browser.Dispose();
            }
        }

        private void ExploreButton_Click(object sender, EventArgs e)
        {
            var cofd = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Multiselect = true,
                InitialDirectory = General.WORK_FOLDER
            };
            var result = cofd.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                textBox2.Text = string.Join(";", cofd.FileNames.ToArray());
                var ids = new List<int>();
                foreach(var name in cofd.FileNames)
                {
                    var dirs = Directory.GetDirectories(name);
                    foreach (var dir in dirs)
                    {
                        var advName = new DirectoryInfo(dir).Name.Split('_')[0];
                        if (int.TryParse(advName, out int id))
                        {
                            ids.Add(id);
                        }
                        var idsArr = ids.ToArray();
                        try
                        {
                            fromBox.Value = idsArr.Min();
                            _idMin = idsArr.Min();
                            toBox.Value = idsArr.Max();
                            _idMax = idsArr.Max();
                        }
                        catch { }
                        UpdateForm();
                    }
                }
            }
            
            /*using (var fbd = new FolderBrowserDialog())
            {
                
                fbd.SelectedPath = General.WORK_FOLDER;
                var result = fbd.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    textBox2.Text = fbd.SelectedPath;
                    var ids = new List<int>();
                    var dirs = Directory.GetDirectories(textBox2.Text);
                    foreach(var dir in dirs)
                    {
                        var name = new DirectoryInfo(dir).Name.Split('_')[0];
                        if(int.TryParse(name, out int id))
                        {
                            ids.Add(id);
                        }
                    }
                    var idsArr = ids.ToArray();
                    try
                    {
                        fromBox.Value = idsArr.Min();
                        _idMin = idsArr.Min();
                        toBox.Value = idsArr.Max();
                        _idMax = idsArr.Max();
                    }
                    catch { }
                    UpdateForm();
                }
            }*/
        }

        private void UploadButton_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Выберите директорию!");
                return;
            }
            if (fromBox.Value == 0 && toBox.Value == 0 || fromBox.Value > toBox.Value)
            {
                MessageBox.Show("Неверно задан интервал");
                return;
            }
            
            _stopFlag = false;
            var thread = new Thread(UploadAdvs)
            {
                IsBackground = true
            };
            //UploadAdvs();
            thread.Start();
            statusLabel.Text = "Загрузка объявлений...";
            uploadButton.Enabled = false;
            exploreButton.Enabled = false;
            parseButton.Enabled = false;
            exportButton.Enabled = false;
            textBox2.Enabled = false;
            stopButton.Enabled = true;
            fromBox.Enabled = false;
            toBox.Enabled = false;
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ActionsForm_Load(object sender, EventArgs e)
        {
            statusLabel.Text = STATUS_BAR_DEFAULT;
            if(!string.IsNullOrEmpty(textBox2.Text))
            {
                var ids = new List<int>();
                var dirs = Directory.GetDirectories(textBox2.Text);
                foreach (var dir in dirs)
                {
                    var name = new DirectoryInfo(dir).Name.Split('_')[0];
                    if (int.TryParse(name, out int id))
                    {
                        ids.Add(id);
                    }
                }
                var idsArr = ids.ToArray();
                try
                {
                    fromBox.Value = idsArr.Min();
                    _idMin = idsArr.Min();
                    toBox.Value = idsArr.Max();
                    _idMax = idsArr.Max();
                }
                catch { }
            }
            UpdateForm();
        }

        private void ActionsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _stopFlag = true;
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            stopButton.Enabled = false;
            statusLabel.Text = "Остановка...";
            _stopFlag = true;
        }

        private void ParseAdvs()
        {
            if (!_db.Connect())
            {
                MessageBox.Show(General.ERROR_DB_CONNECT);
                Invoke(new Action(() =>
                {
                    statusLabel.Text = STATUS_BAR_DEFAULT;
                    uploadButton.Enabled = true;
                    exploreButton.Enabled = true;
                    parseButton.Enabled = true;
                    exportButton.Enabled = true;
                    textBox2.Enabled = true;
                    fromBox.Enabled = true;
                    toBox.Enabled = true;
                }));
                _db.Close();
                return;
            }

            var targets = textBox2.Text.Split(';');
            var def = (int)(toBox.Value - fromBox.Value) + 1;
            uint i = 0;
            foreach (var target in targets)
            {
                var source = (target.Contains("_arc")) ? 2 : 1;
                var dirs = Directory.GetDirectories(target);
                var pathElements = target.Split('\\');
                var comma = pathElements[pathElements.Length - 1].Split('_');
                var type = comma[2] + '_' + comma[3];
                foreach (var dir in dirs)
                {
                    if (_stopFlag)
                    {
                        break;
                    }
                    var dsp = dir.Split('\\');
                    var desp = dsp[dsp.Length - 2];
                    var idStr = dsp[dsp.Length - 1].Split('_')[0];
                    if (!int.TryParse(idStr, out int id))
                    {
                        i++;
                        continue;
                    }
                    if (id < fromBox.Value || id > toBox.Value)
                    {
                        i++;
                        continue;
                    }
                    char status = '-';
                    var singleName = dir.Split('\\')[dir.Split('\\').Length - 1];
                    var fileName = target + "\\" + singleName + ".html";
                    if (!File.Exists(fileName))
                    {
                        fileName = target + "\\" + singleName + "_arch.html";
                        if (!File.Exists(fileName))
                        {
                            fileName = target + "\\" + singleName + "_udalen.html";
                            if (!File.Exists(fileName))
                            {
                                i++;
                                continue;
                            }
                        }
                    }                    
                    if (fileName.Contains("udalen"))
                    {
                        status = 'u';
                        _db.UpdateStatus(id, status, source);
                        continue;
                    }                   

                    FileInfo fileInfo = new FileInfo(fileName);
                    string newname = fileName;
                    if (fileInfo.Length <= 25600)
                    {
                        //if(!fileName.Contains("udalen"))
                        newname = target + "\\" + singleName + "_udalen.html";
                        File.Move(fileName, newname);
                        status = 'u';
                        _db.UpdateStatus(id, status, source);
                        continue;
                    }
                    if (fileName.Contains("arch")) status = 'a';
                    _db.UpdateStatus(id, status, source);

                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(File.ReadAllText(fileName));
                    var restrictions = General.ExtractValueFromDoc(SiteType.Farpost, "restrictions", doc);
                    var agency = General.ExtractValueFromDoc(SiteType.Farpost, "agency", doc);
                    var rooms = General.ExtractValueFromDoc(SiteType.Farpost, "rooms", doc);
                    var description = General.ExtractValueFromDoc(SiteType.Farpost, "description", doc);
                    _db.UpdateAdvData(id, restrictions, string.Empty, agency, rooms, description);

                    var added = General.ExtractValueFromDoc(SiteType.Farpost, "added", doc);
                    if (!String.IsNullOrEmpty(added))
                    {
                        var underlines = singleName.Split('_');
                        /*var savedDateTime = DateTime.MinValue;
                        DateTime.TryParseExact(underlines[2], "yyyyddMM", new CultureInfo("ru-RU"), DateTimeStyles.None, out savedDateTime);*/
                        var savedDateTime = File.GetCreationTime(fileName);
                        var hour = 0;
                        var minute = 0;
                        var dateTime = DateTime.MinValue;
                        var addedArr = added.Split(',');
                        if (addedArr.Length >= 2)
                        {
                            var time = addedArr[0].Trim();
                            var timeArr = time.Split(':');
                            if (timeArr.Length >= 2)
                            {
                                int.TryParse(timeArr[0], out hour);
                                int.TryParse(timeArr[1], out minute);
                            }
                            var date = addedArr[1].Trim().ToLower();
                            if (date == "сегодня")
                            {
                                dateTime = new DateTime(savedDateTime.Year, savedDateTime.Month, savedDateTime.Day, hour, minute, 0);
                                date = "Сегодня";
                            }
                            else if (date == "вчера")
                            {
                                dateTime = new DateTime(savedDateTime.Year, savedDateTime.Month, savedDateTime.Day, hour, minute, 0);
                                dateTime.AddDays(-1);
                                date = "Вчера";
                            }
                            else
                            {
                                var dateArr = date.Split(' ');
                                var format = dateArr.Length == 2 ? "d MMMM" : "d MMMM yyyy";
                                DateTime.TryParseExact(date, format, new CultureInfo("ru-RU"), DateTimeStyles.None, out dateTime);
                            }
                            _db.UpdateDate(id, dateTime, date);
                        }
                    }

                    var username = General.ExtractValueFromDoc(SiteType.Farpost, "username", doc);
                    var phone = General.ExtractValueFromDoc(SiteType.Farpost, "phone", doc);
                    var email = General.ExtractValueFromDoc(SiteType.Farpost, "email", doc);
                    var advsStr = General.GetNumbers(General.ExtractValueFromDoc(SiteType.Farpost, "advs", doc));
                    int.TryParse(advsStr, out int advs);
                    _db.UpdateUserInfo(id, username, phone, email, advs);

                    var address = General.ExtractValueFromDoc(SiteType.Farpost, "address", doc);
                    var hood = General.ExtractValueFromDoc(SiteType.Farpost, "district", doc);
                    var cadastre = General.ExtractValueFromDoc(SiteType.Farpost, "cadastreNumber", doc);
                    var city = General.ExtractValueFromDoc(SiteType.Farpost, "city", doc);
                    _db.UpdateAddr(id, address, string.Empty, string.Empty, string.Empty, string.Empty, cadastre, city);

                    var advName = fileName.Split('\\')[fileName.Split('\\').Length - 1];
                    var advNumber = advName.Split('_')[1];

                    var ID = _db.GetIdByNumber(advNumber);
                    if (ID != -1) // Если объява с таким внутренним идом уже существует
                    {
                        
                        var adv = _db.GetAdvPdata(ID);
                        var advOld = _db.GetAdvPdata(id);
                        if (( General.GetNumbers(adv.Price) != General.GetNumbers(advOld.Price) || General.GetNumbers(adv.Square) != General.GetNumbers(advOld.Square)))
                        {
                            _db.UpdateIsNew(id, 0);                           
                        }
                    }
                    else
                        _db.UpdateIsNew(id, 1);

                    if (type == "КН_ар")
                    {
                        var url = _db.GetUrlById(id);
                        var comType = "n/a";
                        if (url.Contains("/market/"))
                        {
                            comType = "Торговая точка";
                        }
                        else if (url.Contains("/terminal/"))
                        {
                            comType = "Складское помещение";
                        }
                        else if (url.Contains("/workroom/"))
                        {
                            comType = "Производственное помещение";
                        }
                        else if (url.Contains("/office/"))
                        {
                            comType = "Офисное помещенине";
                        }
                        _db.UpdateComTypeSec(advNumber, comType);
                        try { stopButton.Enabled = false; } catch { }
                    }
                    else if (type == "КН_пр" && !string.IsNullOrEmpty(rooms))
                    {
                        _db.UpdateComTypeSec(advNumber, char.ToUpper(rooms[0]) + rooms.Substring(1));
                    }

                    Invoke(new Action(() =>
                    {
                        statusLabel.Text = "Обработка объявления #" + id + " [" + i + "/" + def + "]";
                    }));
                    i++;
                }
            }
            Invoke(new Action(() =>
            {
                statusLabel.Text = STATUS_BAR_DEFAULT;
                uploadButton.Enabled = true;
                exploreButton.Enabled = true;
                parseButton.Enabled = true;
                exportButton.Enabled = true;
                textBox2.Enabled = true;
                fromBox.Enabled = true;
                toBox.Enabled = true;
                stopButton.Enabled = false;
            }));
            _db.Close();
        }


        private void ParseButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Выберите директорию!");
                return;
            }
            if (fromBox.Value == 0 && toBox.Value == 0 || fromBox.Value > toBox.Value)
            {
                MessageBox.Show("Неверно задан интервал");
                return;
            }
            _stopFlag = false;
            var thread = new Thread(ParseAdvs)
            {
                IsBackground = true
            };
            thread.Start();
            //ParseAdvs();
            statusLabel.Text = "Загрузка объявлений...";
            uploadButton.Enabled = false;
            exploreButton.Enabled = false;
            parseButton.Enabled = false;
            exportButton.Enabled = false;
            textBox2.Enabled = false;
            stopButton.Enabled = true;
            fromBox.Enabled = false;
            toBox.Enabled = false;
        }

        private void ExportAdvs()
        {
            if (!_db.Connect())
            {
                MessageBox.Show(General.ERROR_DB_CONNECT);
                Invoke(new Action(() =>
                {
                    statusLabel.Text = STATUS_BAR_DEFAULT;
                    uploadButton.Enabled = true;
                    exploreButton.Enabled = true;
                    parseButton.Enabled = true;
                    exportButton.Enabled = true;
                    textBox2.Enabled = true;
                    fromBox.Enabled = true;
                    toBox.Enabled = true;
                }));
                _db.Close();
                return;
            }
           // try
           // {
                //var taget = textBox2.Text;
                //var dirs = Directory.GetDirectories(taget);
                //var name = new DirectoryInfo(taget).Name + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var targets = textBox2.Text.Split(';');
                var name = string.Empty;
                foreach(var target in targets)
                {
                    name += new DirectoryInfo(target).Name;
                }
                name += '_' + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var fileName = General.WORK_FOLDER + "00_obyav_Farpost_01_Exp-Excel_\\" + name + ".xlsx";
                try
                {
                    if (File.Exists(fileName))
                    {
                        File.Delete(fileName);
                    }
                }
                catch { MessageBox.Show("Не удалось перезаписать excel файл!\nВозможно перезаписываемый файл запущен, закройте его либо удалите файл вручную"); return; }
                int[][] indicies = new int[1][];
                indicies[0] = new int[] { (int)fromBox.Value, (int)toBox.Value };
                string TYPE = "null";
            if (textBox2.Text == "КН_пр" || textBox2.Text == "КН_ар" || textBox2.Text == "ЗУ_пр" || textBox2.Text == "ЗУ_ар" || textBox2.Text == "КВ_пр" || textBox2.Text == "КВ_ар" ||
            textBox2.Text == "ГАР_пр" || textBox2.Text == "ГАР_ар" || textBox2.Text == "ДМ_пр" || textBox2.Text == "ДМ_ар" || textBox2.Text == "ДЧ_пр") { TYPE = textBox2.Text; }                
                //MessageBox.Show(TYPE);
                _db.ExportToExcel((int)fromBox.Value, (int)toBox.Value, fileName, TYPE);
                Invoke(new Action(() =>
                {
                    statusLabel.Text = STATUS_BAR_DEFAULT;
                    uploadButton.Enabled = true;
                    exploreButton.Enabled = true;
                    parseButton.Enabled = true;
                    exportButton.Enabled = true;
                    textBox2.Enabled = true;
                    fromBox.Enabled = true;
                    toBox.Enabled = true;
                }));
                _db.Close();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Что-то пошло нет так!\n" + ex.Message);
            //}
        }

        private void FromBox_ValueChanged(object sender, EventArgs e)
        {
            if (_idMin != 0)
            {
                if (fromBox.Value < _idMin) fromBox.Value = _idMin;   
            }
            UpdateForm();
        }

        private void ToBox_ValueChanged(object sender, EventArgs e)
        {
            if (_idMax != 0)
            {               
                if (toBox.Value > _idMax) toBox.Value = _idMax;
            }
            UpdateForm();
        }

        private void UpdateForm()
        {
            var delta = toBox.Value == 0 && fromBox.Value == 0 ? 0 : toBox.Value - fromBox.Value + 1;
            label3.Text = "Кол-во: " + delta;
            if (delta <= 0)
            {
                uploadButton.Enabled = false;
                parseButton.Enabled = false;
                exportButton.Enabled = false;
            }
            else
            {
                uploadButton.Enabled = true;
                parseButton.Enabled = true;
                exportButton.Enabled = true;
            }
        }

        private void FixButton_Click(object sender, EventArgs e)
        {
            try
            {
                _db.Connect();
                var dirs = Directory.GetDirectories(textBox2.Text);
                foreach (var dir in dirs)
                {
                    string name = dir.Split('\\').Last();
                    if (name.Contains("_0_"))
                    {
                        var temp = dir.Replace(name, "");
                        var num = "_" + _db.GetNumberById(Int32.Parse(name.Split('_')[0])) + "_";
                        name = name.Replace("_0_", num);
                        var newPath = temp + name;            
                        Directory.Move(dir, newPath);                        
                    }
                }
                var files = Directory.GetFiles(textBox2.Text);
                foreach (var file in files)
                {                    
                    var name = file.Split('\\').Last();
                    if (name.Contains("L_")) continue;
                    //if (name.Contains("_0_"))
                    //{
                    var temp = file.Replace(name, "");
                    var id = Int32.Parse(name.Split('_')[0]);
                    var num = "_" + _db.GetNumberById(id) + "_";
                    name = name.Replace("_0_", num);
                    var newPath = temp + name;
                    File.Move(file, newPath);
                    _db.UpdateLinks(id, newPath);
                    //}
                }
            }
            
            catch(Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text == "КН_пр" || textBox2.Text == "КН_ар" || textBox2.Text == "ЗУ_пр" || textBox2.Text == "ЗУ_ар" || textBox2.Text == "КВ_пр" || textBox2.Text == "КВ_ар" ||
           textBox2.Text == "ГАР_пр" || textBox2.Text == "ГАР_ар" || textBox2.Text == "ДМ_пр" || textBox2.Text == "ДМ_ар" || textBox2.Text == "ДЧ_пр")
            {
                exportButton.Enabled = true;
                toBox.Value = 1;
            }
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Выберите директорию!");
                return;
            }
            if (fromBox.Value == 0 && toBox.Value == 0 || fromBox.Value > toBox.Value)
            {
                MessageBox.Show("Неверно задан интервал");
                return;
            }
            _stopFlag = false;
            var thread = new Thread(ExportAdvs)
            {
                IsBackground = true
            };
            thread.Start();
            statusLabel.Text = "Экспорт объявлений...";
            uploadButton.Enabled = false;
            exploreButton.Enabled = false;
            parseButton.Enabled = false;
            exportButton.Enabled = false;
            textBox2.Enabled = false;
            stopButton.Enabled = false;
            fromBox.Enabled = false;
            toBox.Enabled = false;
        }
    }
}
