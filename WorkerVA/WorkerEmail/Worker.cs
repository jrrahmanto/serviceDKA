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
using Google.Apis.Drive.v3;
using Google.Apis.Util.Store;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Spire.Email.IMap;
using Spire.Email;
using OpenPop.Pop3;
using OpenPop.Mime;

namespace WorkerEmail
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private const string PathToCredentials = @"D:\sourcecode\WorkerGoogleDrive\WorkerEmail\bin\Debug\netcoreapp3.1\credentials.json";
        //"E:\ServiceDariTimahDBMS\ServiceVA"
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
                    IronXL.License.LicenseKey = "IRONXL.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-E8D7155B28-DDUBZAO2CK6SZS6-NAGUHRNBVLNI-FJUITVDBUOBQ-3XCIXF7ITTXJ-7W7ND3MR2RG5-K24FCU-LNCDCWWFX2WIEA-PROFESSIONAL.SUB-2GOTUI.RENEW.SUPPORT.13.DEC.2022";
                    List<data_csv> dataCsv = new List<data_csv>();
                    string path = AppDomain.CurrentDomain.BaseDirectory + "\\{0}";
                    //=========connect ke email===========//
                    Pop3Client client = new Pop3Client();
                    client.Connect("outlook.office365.com", 995, true);
                    client.Authenticate("automatic_ptkbi@outlook.com", "Jakarta2021");
                    var messageCount = client.GetMessageCount();

                    for (int j = 0; j < (messageCount); j++)
                    {
                        Message getMessage = client.GetMessage(j + 1);
                        var headers = getMessage.Headers;
                        if (headers.Subject.ToString().ToLower().Contains("mutasi bri, bca dan mandiri"))
                        {
                            monitoringServices("DKA_ServiceVA", "Service untuk convert mutasi BCA, MANDIRI, BRI dari email", "10.10.10.99", "Live");
                            foreach (var attachment in getMessage.FindAllAttachments())
                            {
                                var caption = attachment.ContentType.Name;
                                //kalau mandiri
                                if (caption.Contains("Acc_Statement"))
                                {

                                    string ext = Path.GetExtension(attachment.ContentType.Name);
                                    string truefalse = "False";
                                    //string path = @"D:\AutoSettlementPrice\AutoSettlementPrice\bin\Debug\Setting";

                                    string path_file1 = string.Format(path, "Mandiri" + ext);

                                    if (System.IO.File.Exists(path_file1))
                                    {
                                        System.IO.File.Delete(path_file1);
                                    }

                                    FileStream Stream = new FileStream(path_file1, FileMode.Create);
                                    BinaryWriter BinaryStream = new BinaryWriter(Stream);
                                    BinaryStream.Write(attachment.Body);
                                    BinaryStream.Close();

                                    string[] read_file1 = File.ReadAllLines(path_file1);
                                    for (int i = 1; i < read_file1.Length; i++)
                                    {
                                        string[] dataFile1 = read_file1[i].Split(';');
                                        var nominal = (dataFile1[5].ToString().Replace(".", ","));
                                        //Kalau Mandiri
                                        if (dataFile1[3].Contains("8881020"))
                                        {
                                            if (dataFile1[3].ToLower().Contains("ubp"))
                                            {
                                                truefalse = "True";
                                            }
                                            dataCsv.Add(new data_csv
                                            {
                                                AccountNo = (dataFile1[0].ToString()),
                                                Date = (dataFile1[2].ToString().Remove(dataFile1[2].Length - 9)),
                                                ValDate = (dataFile1[2].ToString()),
                                                Credit = nominal,
                                                Description1 = (dataFile1[3].ToString()),
                                                keterangan = "Valbury",
                                                truefalse = truefalse
                                            });
                                        }
                                        else
                                        {
                                            if (dataFile1[3].ToLower().Contains("ubp"))
                                            {
                                                truefalse = "True";
                                            }
                                            dataCsv.Add(new data_csv
                                            {
                                                AccountNo = (dataFile1[0].ToString()),
                                                Date = (dataFile1[2].ToString().Remove(dataFile1[2].Length - 9)),
                                                ValDate = (dataFile1[2].ToString()),
                                                Credit = nominal,
                                                Description1 = (dataFile1[3].ToString()),
                                                keterangan = "Indodax",
                                                truefalse = truefalse
                                            });
                                        }
                                        truefalse = "False";
                                    }
                                    //write to excell
                                    StringBuilder sb = new StringBuilder();
                                    sb.AppendLine("NO.;ACCOUNT NO;DATE;VAL DATE;CREDIT;DESCRIPTION;KETERANGAN;TRUE OR FALSE");
                                    int k = 1;
                                    foreach (var item in dataCsv)
                                    {
                                        string line = k + ";" + item.AccountNo + ";" + item.Date + ";" + item.ValDate + ";" + item.Credit + ";" + item.Description1 + ";" + item.keterangan + ";" + item.truefalse;
                                        k++;
                                        sb.AppendLine(line);
                                    }
                                    System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "\\Mandiri.csv", sb.ToString());

                                    WorkBook workbook = WorkBook.LoadCSV(AppDomain.CurrentDomain.BaseDirectory + "\\Mandiri.csv", fileFormat: ExcelFileFormat.XLSX, ListDelimiter: ";");
                                    WorkSheet ws = workbook.DefaultWorkSheet;

                                    workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "\\Mandiri.xlsx");

                                    sendFileTelegram("-778112650", AppDomain.CurrentDomain.BaseDirectory + "\\Mandiri.xlsx");

                                    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                                    _logger.LogInformation("Sukses");

                                }
                                //Kalau BCA
                                if (caption.Contains("CorpAcctTrxn"))
                                {
                                    string ext = Path.GetExtension(attachment.ContentType.Name);
                                    string truefalse = "False";
                                    //string path = @"D:\AutoSettlementPrice\AutoSettlementPrice\bin\Debug\Setting";

                                    string path_file1 = string.Format(path, "BCA" + ext);

                                    if (System.IO.File.Exists(path_file1))
                                    {
                                        System.IO.File.Delete(path_file1);
                                    }

                                    FileStream Stream = new FileStream(path_file1, FileMode.Create);
                                    BinaryWriter BinaryStream = new BinaryWriter(Stream);
                                    BinaryStream.Write(attachment.Body);
                                    BinaryStream.Close();

                                    string[] read_file1 = File.ReadAllLines(path_file1);
                                    for (int i = 7; i < read_file1.Length; i++)
                                    {
                                        Console.WriteLine(i);
                                        string[] dataFile1 = read_file1[i].Split('"');
                                        if (dataFile1[0].ToLower().Contains("saldo"))
                                        {
                                            //noting
                                            break;
                                        }
                                        else
                                        {
                                            string[] detail = dataFile1[11].Split(' ');
                                            if (dataFile1[3].ToLower().Contains("kbipayout"))
                                            {
                                                try
                                                {
                                                    if (dataFile1[3].ToLower().Contains("fund transfer"))
                                                    {
                                                        truefalse = "True";
                                                    }
                                                    dataCsv.Add(new data_csv
                                                    {
                                                        Date = (dataFile1[1].ToString() + "/" + DateTime.Now.Year),
                                                        Credit = (detail[0].ToString()),
                                                        Description1 = (dataFile1[3].ToString()),
                                                        Description2 = detail[1],
                                                        keterangan = "Valbury",
                                                        truefalse = truefalse
                                                    });
                                                }
                                                catch (Exception x)
                                                {
                                                    var index = i;
                                                }
                                                
                                            }
                                            else if (dataFile1[3].ToLower().Contains("switching"))
                                            {
                                                try
                                                {
                                                    truefalse = "True";
                                                    dataCsv.Add(new data_csv
                                                    {
                                                        Date = (dataFile1[1].ToString() + "/" + DateTime.Now.Year),
                                                        Credit = (detail[0].ToString()),
                                                        Description1 = (dataFile1[3].ToString()),
                                                        Description2 = detail[1],
                                                        keterangan = "Valbury",
                                                        truefalse = truefalse
                                                    });
                                                }
                                                catch (Exception x)
                                                {
                                                    var index = i;
                                                }

                                            }
                                            else if (dataFile1[3].ToLower().Contains("otomatis"))
                                            {
                                                try
                                                {
                                                    dataCsv.Add(new data_csv
                                                    {
                                                        Date = (dataFile1[1].ToString() + "/" + DateTime.Now.Year),
                                                        Credit = (detail[0].ToString()),
                                                        Description1 = (dataFile1[3].ToString()),
                                                        Description2 = detail[1],
                                                        keterangan = "BCA",
                                                        truefalse = truefalse
                                                    });
                                                }
                                                catch (Exception x)
                                                {
                                                    var index = i;
                                                }

                                            }
                                            else
                                            {
                                                try
                                                {
                                                    if (dataFile1[3].ToLower().Contains("fund transfer"))
                                                    {
                                                        truefalse = "True";
                                                    }
                                                    dataCsv.Add(new data_csv
                                                    {
                                                        Date = (dataFile1[1].ToString() + "/" + DateTime.Now.Year),
                                                        Credit = (detail[0].ToString()),
                                                        Description1 = (dataFile1[3].ToString()),
                                                        Description2 = detail[1],
                                                        keterangan = "Indodax",
                                                        truefalse = truefalse
                                                    });
                                                }
                                                catch (Exception x)
                                                {
                                                    var index = i;
                                                }
                                               
                                            }
                                            truefalse = "False";
                                        }
                                    }
                                    //write to excell
                                    StringBuilder sb = new StringBuilder();
                                    sb.AppendLine("NO.;DATE;CREDIT;DESCRIPTION 1;DESCRIPTION 2;KETERANGAN;TRUE OR FALSE");
                                    int k = 1;
                                    foreach (var item in dataCsv)
                                    {
                                        string line = k + ";" + item.Date + ";" + item.Credit + ";" + item.Description1 + ";" + item.Description2 + ";" + item.keterangan + ";" + item.truefalse;
                                        k++;
                                        sb.AppendLine(line);
                                    }
                                    System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "\\BCA.csv", sb.ToString());
                                    //IronXL.License.LicenseKey = "IRONXL.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-E8D7155B28-DDUBZAO2CK6SZS6-NAGUHRNBVLNI-FJUITVDBUOBQ-3XCIXF7ITTXJ-7W7ND3MR2RG5-K24FCU-LNCDCWWFX2WIEA-PROFESSIONAL.SUB-2GOTUI.RENEW.SUPPORT.13.DEC.2022";
                                    WorkBook workbook = WorkBook.LoadCSV(AppDomain.CurrentDomain.BaseDirectory + "\\BCA.csv", fileFormat: ExcelFileFormat.XLSX, ListDelimiter: ";");
                                    WorkSheet ws = workbook.DefaultWorkSheet;

                                    workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "\\BCA.xlsx");

                                    sendFileTelegram("-778112650", AppDomain.CurrentDomain.BaseDirectory + "\\BCA.xlsx");
                                    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                                    _logger.LogInformation("Sukses");

                                }
                                //Kalau BRI
                                if (caption.Contains("DD_ONLINE"))
                                {
                                    string ext = Path.GetExtension(attachment.ContentType.Name);
                                    string truefalse = "False";
                                    //string path = @"D:\AutoSettlementPrice\AutoSettlementPrice\bin\Debug\Setting";

                                    string path_file1 = string.Format(path, "BRI" + ext);

                                    if (System.IO.File.Exists(path_file1))
                                    {
                                        System.IO.File.Delete(path_file1);
                                    }

                                    FileStream Stream = new FileStream(path_file1, FileMode.Create);
                                    BinaryWriter BinaryStream = new BinaryWriter(Stream);
                                    BinaryStream.Write(attachment.Body);
                                    BinaryStream.Close();

                                    //extract data
                                    List<data_bri> dataBRI = new List<data_bri>();
                                    WorkBook workbook = WorkBook.Load(path_file1);
                                    WorkSheet sheet = workbook.WorkSheets.First();
                                    //string x = sheet["A:P"].ToString();
                                    //Select cells easily in Excel notation and return the calculated value
                                    string[] data = sheet["A:P"].ToString().Split("\r\n");
                                    for (int i = 13; i < data.Length; i++)
                                    {
                                        if (data[i] == "\t\t\t\t\t\t\t\t\t\t\t")
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            string[] newdata = data[i].Split("\t");
                                            if (newdata[2].Contains("13042"))
                                            {
                                                if (newdata[2].ToLower().Contains("atm") || newdata[2].ToLower().Contains("briva"))
                                                {
                                                    truefalse = "True";
                                                }
                                                dataBRI.Add(new data_bri
                                                {
                                                    Date = (newdata[0].ToString()),
                                                    Time = (newdata[1].ToString()),
                                                    Remark = (newdata[2].ToString().Replace(";", ":")),
                                                    Debet = (newdata[6].ToString()),
                                                    Credit = (newdata[9].ToString()),
                                                    keterangan = "INDODAX",
                                                    truefalse = truefalse
                                                });
                                            }
                                            else if (newdata[2].Contains("13275"))
                                            {
                                                if (newdata[2].ToLower().Contains("atm") || newdata[2].ToLower().Contains("briva"))
                                                {
                                                    truefalse = "True";
                                                }
                                                dataBRI.Add(new data_bri
                                                {
                                                    Date = (newdata[0].ToString()),
                                                    Time = (newdata[1].ToString()),
                                                    Remark = (newdata[2].ToString().Replace(";", ":")),
                                                    Debet = (newdata[6].ToString()),
                                                    Credit = (newdata[9].ToString()),
                                                    keterangan = "VALBURY",
                                                    truefalse = truefalse
                                                });
                                            }
                                            else if (newdata[2].Contains("12362"))
                                            {
                                                if (newdata[2].ToLower().Contains("atm") || newdata[2].ToLower().Contains("briva"))
                                                {
                                                    truefalse = "True";
                                                }
                                                dataBRI.Add(new data_bri
                                                {
                                                    Date = (newdata[0].ToString()),
                                                    Time = (newdata[1].ToString()),
                                                    Remark = (newdata[2].ToString().Replace(";", ":")),
                                                    Debet = (newdata[6].ToString()),
                                                    Credit = (newdata[9].ToString()),
                                                    keterangan = "VALBURY",
                                                    truefalse = truefalse
                                                });
                                            }
                                            else
                                            {
                                                if (newdata[2].ToLower().Contains("atm") || newdata[2].ToLower().Contains("briva"))
                                                {
                                                    truefalse = "True";
                                                }
                                                dataBRI.Add(new data_bri
                                                {
                                                    Date = (newdata[0].ToString()),
                                                    Time = (newdata[1].ToString()),
                                                    Remark = (newdata[2].ToString().Replace(";", ":")),
                                                    Debet = (newdata[6].ToString()),
                                                    Credit = (newdata[9].ToString()),
                                                    keterangan = "-",
                                                    truefalse = truefalse
                                                });
                                            }
                                            truefalse = "False";
                                        }
                                    }

                                    //write to excell
                                    StringBuilder sb = new StringBuilder();
                                    sb.AppendLine("NO;DATE;TIME;REMARK;DEBET;CREDIT;KETERANGAN;TRUE OR FALSE");
                                    int k = 1;
                                    foreach (var item in dataBRI)
                                    {
                                        string line = k + ";" + item.Date + ";" + item.Time + ";" + item.Remark + ";" + item.Debet + ";" + item.Credit + ";" + item.keterangan + ";" + item.truefalse;
                                        k++;
                                        sb.AppendLine(line);
                                    }
                                    System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "\\BRI_CSV.csv", sb.ToString());
                                    WorkBook wb = WorkBook.LoadCSV(AppDomain.CurrentDomain.BaseDirectory + "\\BRI_CSV.csv", fileFormat: ExcelFileFormat.XLSX, ListDelimiter: ";");
                                    WorkSheet ws = wb.DefaultWorkSheet;

                                    wb.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "\\BRI_Result.xlsx");

                                    sendFileTelegram("-778112650", AppDomain.CurrentDomain.BaseDirectory + "\\BRI_Result.xlsx");

                                    //string uploadedFileId = await uploadToGoogleDrive(AppDomain.CurrentDomain.BaseDirectory + "\\result1.xlsx");
                                    //SendMessage("6281310215750-1621573569@g.us", "Processing finish at : " + DateTime.Now.ToString("HH:mm:ss") + " https://drive.google.com/file/d/" + uploadedFileId + "/view?usp=drivesdk");
                                    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                                    _logger.LogInformation("Sukses");

                                }
                            }
                            client.DeleteMessage(j + 1);
                        }
                    }
                    client.Disconnect();
                }

                catch (Exception ex)
                {
                    //SendMessage("6281310215750-1621573569@g.us", "Extract fail : "+ex.Message);
                    //monitoringServices("DKA_ServiceVA", "Service untuk convert mutasi BCA, MANDIRI, BRI dari email", "10.10.10.99", ex.Message);
                    _logger.LogInformation("Worker Eror : " + ex.Message + " {time}", DateTimeOffset.Now);

                }

                await Task.Delay(15000, stoppingToken);
            }
        }
        public class data_csv
        {
            public string AccountNo { get; set; }
            public string Date { get; set; }
            public string ValDate { get; set; }
            public string TransactionCode { get; set; }
            public string Description1 { get; set; }
            public string Description2 { get; set; }
            public string truefalse { get; set; }
            public string Debit { get; set; }
            public string Credit { get; set; }
            public string keterangan { get; set; }

        }
        public class data_bri
        {
            public string AccountNo { get; set; }
            public string Date { get; set; }
            public string Time { get; set; }
            public string Remark { get; set; }
            public string Debet { get; set; }
            public string Credit { get; set; }
            public string Ledger { get; set; }
            public string truefalse { get; set; }
            public string keterangan { get; set; }

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
        private static string monitoringServices(string servicename, string servicedescription, string servicelocation, string appstatus)
        {
            string jsonString = "{" +
                                "\"service_name\" : \"" + servicename + "\"," +
                                "\"service_description\": \"" + servicedescription + "\"," +
                                "\"service_location\":\"" + servicelocation + "\"," +
                                "\"app_status\":\""+ appstatus + "\","+
                                "}";
            var client = new RestClient("http://10.10.10.99:84/api/ServiceStatus");

            RestRequest requestWa = new RestRequest("http://10.10.10.99:84/api/ServiceStatus", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("data", jsonString);
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
        private async Task<string> uploadToGoogleDrive(string path)
        {
            string uploadedFileId;
            try
            {
                var token = new FileDataStore("UserCredentialStoragePath", true);
                UserCredential credential;
                string[] scopes = new string[] { DriveService.Scope.Drive, DriveService.Scope.DriveFile, };
                await using (var stream = new FileStream(PathToCredentials, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        scopes,
                        "userName",
                        CancellationToken.None,
                        new FileDataStore("AwsomeAooToken")
                        ).Result;
                }
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "AwsomeAoo"
                });

                var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = "report.xlsx",
                    Parents = new List<string> { "1QujweILVyQ2hQAe5-LHZ2hftwTF9H1HJ" }
                };

                await using (var fssource = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    var request = service.Files.Create(fileMetadata, fssource, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    request.Fields = "*";
                    var result = await request.UploadAsync(CancellationToken.None);
                    if (result.Status == Google.Apis.Upload.UploadStatus.Failed)
                    {
                        Console.WriteLine($"Error uploading file: {result.Exception.Message}");
                    }
                    uploadedFileId = request.ResponseBody.Id;
                }
                return uploadedFileId;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return ex.Message;
            }
        }
    }
}
