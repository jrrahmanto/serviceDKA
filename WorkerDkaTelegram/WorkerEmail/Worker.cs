using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WDSE;
using WDSE.Decorators;
using WDSE.ScreenshotMaker;
using System.Drawing;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json;
using RestSharp;
using System.Net;
using Microsoft.Graph;
using System.Text;
using IronXL;
using Google.Apis.Util.Store;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using System.Data.SqlClient;
using System.Configuration;
using Telegram.Bot;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Args;
using MySql.Data.MySqlClient;
using System.Data;
using Microsoft.Reporting.WebForms;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Net.Mime;
using System.IO.Compression;
using System.Net.Mail;

namespace WorkerEmail
{
    public class Worker : BackgroundService
    {
        //tutorial google drive
        //https://www.youtube.com/watch?v=pHOweM1Gl6c
        //create project
        //open api drive

        private readonly ILogger<Worker> _logger;
        static TelegramBotClient Bot = new TelegramBotClient("5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4");
        public static String connectionString = "Data Source=KBIDRC-TIMAH-DBMS.ptkbi.com;Initial Catalog=TIN_KBI;Persist Security Info=True;User ID=dbapp;Password=P@ssw0rd2022";
        private const string PathToCredentials = @"E:\ServiceDariTimahDBMS\ServiceTINOperationalTelegram\credentials.json";

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }
        public override Task StartAsync(CancellationToken cancellationToken)
        {
            return base.StartAsync(cancellationToken);
        }
        public override Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service stopped");
            return base.StopAsync(cancellationToken);
        }
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {

            try
            {
                Bot.StartReceiving(Array.Empty<UpdateType>());
                Bot.OnUpdate += BotOnUpdateReceived;
            }
            catch (Exception x)
            {
                //SendMessage("6289630870658@c.us", "Tradeprogress fail " + x.Message);
            }
        }
        private static async void BotOnUpdateReceived(object sender, UpdateEventArgs e)
        {

            var message = e.Update.Message;

            if (message.Type == MessageType.Text)
            {
                var text = message.Text;

                if (message.ReplyToMessage != null)
                {
                    var msg = message.ReplyToMessage.Text;
                    //filter untuk approve shipping instructions
                    if (msg.Contains("#report dana jaminan#"))
                    {
                        generateReportDanaJaminan(message.Chat.Id, text);
                    }
                    if (msg.Contains("#send report dana jaminan#"))
                    {
                        sendReportDanaJaminan(message.Chat.Id, text);
                    }
                }
                if (text == "/reportdanajaminan")
                {
                    monitoringServices("DKA_BotTelegram", "generate report dana jaminan", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#report dana jaminan# Please reply business date\n{yyyy-mm-dd}" + "_", ParseMode.Markdown);
                }
                if (text == "/sendreport")
                {
                    monitoringServices("DKA_BotTelegram", "generate report dana jaminan", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#send report dana jaminan# Please reply business date\n{yyyy-mm-dd}" + "_", ParseMode.Markdown);
                }
            }
            if (message.Type == MessageType.Document)
            {
                //download all document
                var file = await Bot.GetFileAsync(message.Document.FileId);
                FileStream fs = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\" + message.Document.FileName, FileMode.Create);
                await Bot.DownloadFileAsync(file.FilePath, fs);
                fs.Close();
                fs.Dispose();
                //filer disini untuk insert bukti tf
                if (message.Document.FileName.ToLower().Contains("dana jaminan"))
                {
                    monitoringServices("DKA_BotTelegram", "upload data dana jaminan", "10.10.10.99", "Live");
                    insertDanaJaminan(message.Chat.Id, fs.Name);
                }
            }
        }
        public static void generateReportDanaJaminan(long chat_id, string msg)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Generate start_", ParseMode.Markdown);

                var dr = new DanaJaminanTableAdapters.DanaJaminanTableAdapter();
                var dt = dr.GetDataByBd(msg);
                List<string> filePaths = new List<string>();
                if (dt.Count != 0)
                {
                    foreach (var item in dt)
                    {
                        var code = item.code;
                        string path = getReportDanaJaminanSSRSWord("RptDanaJaminan", " &businessdate=" + msg + "&code=" + code, item.name, "DanaJaminan");
                        filePaths.Add(path);
                    }
                }
                using (ZipArchive archive = ZipFile.Open(AppDomain.CurrentDomain.BaseDirectory + "report\\" + "Dana Jaminan.rar", ZipArchiveMode.Create))
                {
                    foreach (var fPath in filePaths)
                    {
                        archive.CreateEntryFromFile(fPath, Path.GetFileName(fPath));
                    }
                }
                sendFileTelegram(chat_id.ToString(), AppDomain.CurrentDomain.BaseDirectory + "report\\" + "Dana Jaminan.rar");
                File.Delete(AppDomain.CurrentDomain.BaseDirectory + "report\\" + "Dana Jaminan.rar");
            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Generate Error " + ex.Message + "_", ParseMode.Markdown);
            }
        }
        public static void sendReportDanaJaminan(long chat_id, string msg)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Send report start_", ParseMode.Markdown);

                var dr = new DanaJaminanTableAdapters.DanaJaminanTableAdapter();
                var dt = dr.GetDataByBd(msg);
                List<string> filePaths = new List<string>();
                if (dt.Count != 0)
                {
                    foreach (var item in dt)
                    {
                        var code = item.code;
                        string path = getReportDanaJaminanSSRSWord("RptDanaJaminan", " &businessdate=" + msg + "&code=" + code, item.name, "DanaJaminan");
                        var dt_send = new DanaJaminanTableAdapters.SendEmailTableAdapter();
                        var dr_send = dt_send.GetData(item.code);
                        if (dr_send.Count != 0)
                        {
                            foreach (var item_send in dr_send)
                            {
                                try
                                {
                                    string file = AppDomain.CurrentDomain.BaseDirectory + "\\index.html";
                                    string text = File.ReadAllText(file);
                                    text = text.Replace("#pelaku#", item_send.name);

                                    MailMessage message = new MailMessage();
                                    SmtpClient smtp = new SmtpClient();
                                    message.From = new MailAddress("pb@ptkbi.com");
                                    message.To.Add(new MailAddress(item_send.email));
                                    message.Subject = "Pengiriman report Dana Jaminan";
                                    message.IsBodyHtml = true; //to make message body as html  
                                    message.Body = text;
                                    message.Attachments.Add(new Attachment(path));
                                    smtp.Port = 25;
                                    smtp.Host = "10.10.10.2"; //for gmail host  
                                    smtp.EnableSsl = true;
                                    smtp.UseDefaultCredentials = false;
                                    smtp.EnableSsl = false;
                                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                                    new Task(delegate
                                    {
                                        smtp.Send(message);
                                    }).Start();
                                    WriteToFile(item_send.email + " Success send email");
                                }
                                catch (Exception ex)
                                {
                                    WriteToFile(item_send.email + " Fail send email " +ex.Message);
                                }
                            }
                        }
                    }
                }
                string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
                sendFileTelegram(chat_id.ToString(), filepath);
                File.Delete(filepath);
                File.Delete(AppDomain.CurrentDomain.BaseDirectory + "report\\" + "Dana Jaminan.rar");
            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Generate Error " + ex.Message + "_", ParseMode.Markdown);
            }
        }
        public static void insertDanaJaminan(long chat_id, string path)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Insert Dana Jaminan Start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                IronXL.License.LicenseKey = "IRONXL.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-E8D7155B28-DDUBZAO2CK6SZS6-NAGUHRNBVLNI-FJUITVDBUOBQ-3XCIXF7ITTXJ-7W7ND3MR2RG5-K24FCU-LNCDCWWFX2WIEA-PROFESSIONAL.SUB-2GOTUI.RENEW.SUPPORT.13.DEC.2022";
                WorkBook workbook = WorkBook.Load(path);
                WorkSheet sheet = workbook.WorkSheets.First();
                string[] data = sheet["A:R"].ToString().Split("\r\n");
                List<data_csv> dataCsv = new List<data_csv>();
                for (int i = 0; i < data.Length; i++)
                {
                    string[] newdata = data[i].Split("\t");
                    if (newdata[0] == "NAMA")
                    {

                    }
                    else if (newdata[0] == "")
                    {
                        break;
                    }
                    else
                    {
                        dataCsv.Add(new data_csv
                        {
                            code = newdata[1],
                            bank = newdata[2],
                            jumlah = newdata[5],
                            jangkawaktu = newdata[6],
                            tanggalpenempatan = newdata[7],
                            jatuhtempo = newdata[8],
                            sukubunga = Convert.ToDecimal(newdata[9]) * 100,
                            bungabruto = newdata[10],
                            pph = newdata[11],
                            bunga = newdata[12],
                            admin = newdata[13],
                            transferdana = newdata[15],
                            transferdanakbi = newdata[16],
                            penempatan = newdata[17]
                        });

                    }
                }
                var dr_insert = new DanaJaminanTableAdapters.DanaJaminan1TableAdapter();
                foreach (var item in dataCsv)
                {
                    var dt = dr_insert.GetDataByDate(item.tanggalpenempatan, item.jatuhtempo, item.code, Convert.ToDecimal(item.penempatan));
                    var count = dt.Count();
                    if (count == 0)
                    {
                        dr_insert.Insert(DateTime.Now.Date, item.code, item.bank, Convert.ToDecimal(item.jumlah), Convert.ToInt32(item.jangkawaktu), Convert.ToDateTime(item.tanggalpenempatan), Convert.ToDateTime(item.jatuhtempo), item.sukubunga, Convert.ToDecimal(item.bungabruto), Convert.ToDecimal(item.pph), Convert.ToDecimal(item.bunga), 0, Convert.ToDecimal(item.admin), Convert.ToDecimal(item.transferdana), Convert.ToDecimal(item.transferdanakbi), Convert.ToDecimal(item.penempatan), "T", "T", 1, "1");
                    }
                    else
                    {
                        Bot.SendTextMessageAsync(chat_id, "_Data sudah pernah di input " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                    }
                }
                Bot.SendTextMessageAsync(chat_id, "_Processing success " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Processing upload proof of payment fail " + ex.Message + "\n" + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
            }
        }
        public static string getReportDanaJaminanSSRSWord(string reportname, string param, string filename, string pathreport)
        {
            try
            {
                string url = "http://10.12.5.60/ReportServerEOD?/" + pathreport + "/" + reportname + "&rs:Command=Render&rs:Format=PDF&rc:OutputFormat=PDF" + param;

                System.Net.HttpWebRequest Req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                Req.Credentials = new NetworkCredential("administrator", "Jakarta01");
                Req.Method = "GET";

                string path = AppDomain.CurrentDomain.BaseDirectory + "report\\" + filename + ".pdf";

                System.Net.WebResponse objResponse = Req.GetResponse();
                System.IO.FileStream fs = new System.IO.FileStream(path, System.IO.FileMode.Create);
                System.IO.Stream stream = objResponse.GetResponseStream();

                byte[] buf = new byte[1024];
                int len = stream.Read(buf, 0, 1024);
                while (len > 0)
                {
                    fs.Write(buf, 0, len);
                    len = stream.Read(buf, 0, 1024);
                }
                stream.Close();
                fs.Close();
                return path;

            }
            catch (Exception ex)
            {

                throw;
            }
        }
        private static void sendFileTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;

            var client = new RestClient("https://api.telegram.org/bot5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4/sendDocument");
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4/sendDocument", Method.Post);


            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "multipart/form-data");
            requestWa.AddParameter("chat_id", chatId);
            requestWa.AddFile("document", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);
        }
        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!System.IO.File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = System.IO.File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = System.IO.File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        public class data_csv
        {
            public string businessdate { get; set; }
            public string code { get; set; }
            public string bank { get; set; }
            public string jumlah { get; set; }
            public string jangkawaktu { get; set; }
            public string tanggalpenempatan { get; set; }
            public string jatuhtempo { get; set; }
            public Decimal sukubunga { get; set; }
            public string bungabruto { get; set; }
            public string pph { get; set; }
            public string bunga { get; set; }
            public string adjustment { get; set; }
            public string admin { get; set; }
            public string transferdana { get; set; }
            public string transferdanakbi { get; set; }
            public string penempatan { get; set; }
            public string aro { get; set; }
            public string multiple { get; set; }
            public string sequence { get; set; }
            public string flag { get; set; }


        }
        private static string monitoringServices(string servicename, string servicedescription, string servicelocation, string appstatus)
        {
            string jsonString = "{" +
                                "\"service_name\" : \"" + servicename + "\"," +
                                "\"service_description\": \"" + servicedescription + "\"," +
                                "\"service_location\":\"" + servicelocation + "\"," +
                                "\"app_status\":\"" + appstatus + "\"," +
                                "}";
            var client = new RestClient("http://10.10.10.99:84/api/ServiceStatus");

            RestRequest requestWa = new RestRequest("http://10.10.10.99:84/api/ServiceStatus", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("data", jsonString);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }

    }
}
