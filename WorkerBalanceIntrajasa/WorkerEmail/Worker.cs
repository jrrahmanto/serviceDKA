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
using System.Text.RegularExpressions;
using Telegram.Bot;
using Telegram.Bot.Types.Enums;

namespace WorkerEmail
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        static TelegramBotClient Bot = new TelegramBotClient("5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4");

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
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                    string file = AppDomain.CurrentDomain.BaseDirectory + "\\numberregistered.txt";
                    string[] text = File.ReadAllLines(file);
                    List<string> number = new List<string>();
                    
                    IWebDriver ChromeDriver2 = new ChromeDriver();
                    ChromeDriver2.Manage().Window.Maximize();
                    ChromeDriver2.Navigate().GoToUrl("https://report.intraremittance.com/login.remit");
                    IWebElement username = ChromeDriver2.FindElement(By.Id("username"));
                    IWebElement password = ChromeDriver2.FindElement(By.Id("password"));
                    IWebElement submit_login = ChromeDriver2.FindElement(By.Id("submit"));
                    username.SendKeys("dti-kbi");
                    password.SendKeys("dti1tskbI");
                    submit_login.Click();
                    ChromeDriver2.Navigate().GoToUrl("https://report.intraremittance.com/account!balance.remit");
                    IWebElement table2 = ChromeDriver2.FindElement(By.Id("list-table"));

                    var html_text = table2.GetAttribute("innerHTML");
                    string allText = html_text;//File.ReadAllText(pathHTML);
                    var htmlremove = Regex.Replace(allText, "<.*?>", String.Empty);
                    var data = htmlremove.Split(new[] { "\r\n\t\t\t\t\t\t\t\t" }, StringSplitOptions.None);
                    string idx = data[5].Replace("\r\n\t\t\t\t\t\t\t\r\n\t\t\t\t\t\t\r\n\t\t\t\t\t\t\t", "");
                    string val = data[10].Replace("\r\n\t\t\t\t\t\t\t\r\n\t\t\t\t\t\t\r\n\t\t\t\t\t", "");

                    ChromeDriver2.Navigate().GoToUrl("https://report.intraremittance.com/setting.remit");
                    IWebElement table = ChromeDriver2.FindElement(By.Id("list-table"));

                    html_text = table.GetAttribute("innerHTML");
                    allText = html_text;//File.ReadAllText(pathHTML);
                    htmlremove = Regex.Replace(allText, "<.*?>", String.Empty);
                    data = htmlremove.Split(new[] { "\r\n\t\t\t\t\t\t\t\t" }, StringSplitOptions.None);
                    string bri = data[94] + " : " + data[96];
                    string bca = data[2] + " : " + data[4];
                    string bni = data[48] + " : " + data[50];
                    string mandiri = data[140] + " : " + data[142];
                    ChromeDriver2.Quit();
                    //sendTelegram("-778112650", "Notification Balance IntraRemit\n1. Indodax : " + idx + "\n2. Valbury : " + val + "\n3. "+bri+ "\n4. " + bca + "\n5. " + bni + "\n6. " + mandiri + "\nTimeStamp: " + DateTime.Now.ToString("hh:mm:ss"));
                    sendTelegram("-778112650", "Notification Balance IntraRemit\n1. " + bri + "\n2. " + bca + "\n3. " + bni + "\n4. " + mandiri + "\n5. Indodax : " + idx + "\n6. Valbury : " + val + "\nTimeStamp: " + DateTime.Now.ToString("hh:mm:ss"));
                    for (int i = 0; i < text.Length; i++)
                    {
                        string[] chat_id = text[i].Split(" ");
                        SendMessage(chat_id[0], "1. " + bri + "\n2. " + bca + "\n3. " + bni + "\n4. " + mandiri+"\n5.Indodax : " + idx + "\n6.Valbury : " + val );

                    }

                    _logger.LogInformation("Sukses");
                    monitoringServices("DOP_BalanceIntraRemit","I", "Notifikasi balance dari web intra remit");

                }
                catch (Exception ex)
                {
                    _logger.LogInformation("Worker Eror : " + ex.Message + " {time}", DateTimeOffset.Now);
                    monitoringServices("DOP_BalanceIntraRemit","E", "Notifikasi balance dari web intra remit Eror: " + ex.Message);

                }

                await Task.Delay(3600000, stoppingToken);
            }
        }
        private static void sendFileTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;

            var client = new RestClient("https://api.telegram.org/bot2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U/sendDocument");
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U/sendDocument", Method.Post);


            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "multipart/form-data");
            requestWa.AddParameter("chat_id", chatId);
            requestWa.AddFile("document", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);
        }
        private static string SendFile(string chatId, string data, string caption)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendFile?token=jkdjtwjkwq2gfkac");


            RestRequest requestWa = new RestRequest("https://api.chat-api.com/instance127354/sendFile?token=jkdjtwjkwq2gfkac", Method.Post);

            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("chatId", chatId);
            requestWa.AddParameter("filename", "1.png");
            requestWa.AddParameter("body", data);
            requestWa.AddParameter("caption", caption);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }
        public class MessageChat
        {
            public string id { get; set; }
            public string body { get; set; }
            public string fromMe { get; set; }
            public string self { get; set; }
            public string isForwarded { get; set; }
            public string author { get; set; }
            public double time { get; set; }
            public string chatId { get; set; }
            public int messageNumber { get; set; }
            public string type { get; set; }
            public string senderName { get; set; }
            public string caption { get; set; }
            public string quotedMsgBody { get; set; }
            public string quotedMsgId { get; set; }
            public string quotedMsgType { get; set; }
            public string chatName { get; set; }
        }
        public class ResponseChat
        {
            public IEnumerable<MessageChat> messages { get; set; }
            public int lastMessageNumber { get; set; }
        }
        private static string SendMessage(string chatId, string body)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac");

            RestRequest requestWa = new RestRequest("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("phone", chatId);
            requestWa.AddParameter("body", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }
        private static string GetMessageList(string chatId, int lastMessageNumber)
        {
            string url = "https://api.chat-api.com/instance127354/messages?token=jkdjtwjkwq2gfkac&lastMessageNumber=" + lastMessageNumber + "&chatId=" + chatId;
            var client = new RestClient(url);
            var requestWa = new RestRequest(url, Method.Get);
            var responseWa = client.ExecuteGetAsync(requestWa);
            return responseWa.Result.Content;

        }
        private static void sendTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;

            var client = new RestClient("https://api.telegram.org/bot5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4/sendMessage?chat_id=" + chatId + "&text=" + body);
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot5478187618:AAENfPcaia3OMwc3alj57qil0uN7JrPFPP4/sendMessage?chat_id=" + chatId + "&text=" + body, Method.Get);
            requestWa.Timeout = -1;
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);
        }
        private static string monitoringServices(string servicename, string status, string desc)
        {
            string jsonString = "{" +
                                "\"name\" : \"" + servicename + "\"," +
                                "\"logstatus\": \"" + status + "\"," +
                                "\"logdesc\":\"" + desc + "\"," +
                                "}";
            var client = new RestClient("https://apiservicekbi.azurewebsites.net/api/ServiceStatus");

            RestRequest requestWa = new RestRequest("https://apiservicekbi.azurewebsites.net/api/ServiceStatus", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("data", jsonString);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }

    }
}
