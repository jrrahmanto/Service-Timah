using IronXL;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using RestSharp;
using System.Data;
using System.IO;
using System.Net;
using System.Security.Claims;
using System.Text;

namespace WorkerCheckDataTKS
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                string file = AppDomain.CurrentDomain.BaseDirectory + "\\ConfigNumber.txt";
                try
                {
                    string[] text = File.ReadAllLines(file);
                    foreach (var item in text)
                    {
                        var log = item.Split(" ");
                        var getmessage = await getMessage(log[0], log[1]);
                        if (getmessage != "{\"messages\":[]}")
                        {
                            ResponseChat rc = JsonConvert.DeserializeObject<ResponseChat>(getmessage);
                            if (rc.messages.Count() != 0)
                            {
                                foreach (var data in rc.messages)
                                {
                                    string alltext = File.ReadAllText(file);
                                    string[] textnew = File.ReadAllLines(file);
                                    string newvalue = alltext.Replace(textnew[Convert.ToInt32(log[2])], log[0] + " " + data.messageNumber + " " + log[2]);
                                    File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "\\ConfigNumber.txt", newvalue);
                                    if (data.type == "document")
                                    {
                                        SendMessage(log[0], "Process checking data start " + DateTime.Now.ToString("HH:mm:ss"));

                                        string path = AppDomain.CurrentDomain.BaseDirectory + "\\document\\" + data.caption;
                                        using (WebClient wc = new WebClient())
                                        {
                                            wc.DownloadFile(
                                                new System.Uri(data.body),
                                                path
                                            );
                                        }
                                        string result = await readExcel(path);
                                        SendMessage(log[0], result);
                                        SendMessage(log[0], "Process checking data finish " + DateTime.Now.ToString("HH:mm:ss"));
                                        monitoringServices("DOP_CheckDataTKS", "I", "Process Success");
                                    }
                                    sendTelegram("-1001649045625", "==History chat WA API==\n" + data.body.ToString() + "\n===End===");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    monitoringServices("DOP_CheckDataTKS", "E", "Process Fail "+ex.Message);
                }
                await Task.Delay(10000, stoppingToken);
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss"));
            }
        }

        public async Task<string> getMessage(string number, string lastnumber)
        {
            var url = "https://api.1msg.io/127354/messages?token=jkdjtwjkwq2gfkac&lastMessageNumber=" + lastnumber + "&chatId=" + number;
            var client = new RestClient(url);
            var request = new RestRequest(url, Method.Get);
            request.AddHeader("Content-Type", "application/json");
            //var body = new
            //{
            //    name = "Ajay Kumar",
            //    job = "Developer"
            //};
            //var bodyy = JsonConvert.SerializeObject(body);
            //request.AddBody(bodyy, "application/json");
            RestResponse response = await client.ExecuteAsync(request);
            var output = response.Content;
            return output;
        }
        public class ResponseChat
        {
            public IEnumerable<MessageChat> messages { get; set; }
        }
        public class MessageChat
        {
            public string messageNumber { get; set; }
            public string chatId { get; set; }
            public string type { get; set; }
            public string body { get; set; }
            public string caption { get; set; }
        }
        public static async Task<string> readExcel(string path)
        {
            string depositor = "ok";
            string warehouse = "ok";
            string contract_code = "ok";
            string tradeaccount = "ok";
            string quantity = "ok";

            IronXL.License.LicenseKey = "IRONXL.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-E8D7155B28-DDUBZAO2CK6SZS6-NAGUHRNBVLNI-FJUITVDBUOBQ-3XCIXF7ITTXJ-7W7ND3MR2RG5-K24FCU-LNCDCWWFX2WIEA-PROFESSIONAL.SUB-2GOTUI.RENEW.SUPPORT.13.DEC.2022";
            WorkBook workbook = WorkBook.Load(path);
            WorkSheet sheet = workbook.WorkSheets.First();

            string[] data = sheet["A:Q"].ToString().Split("\r\n");

            for (int i = 4; i < data.Length; i++)
            {
                var splitdata = data[i].Split("\t");
                depositor = await execquery("SELECT * FROM SKD.ClearingMember WHERE Name = '"+ splitdata[3] + "'", splitdata[1]);
                warehouse = await execquery("SELECT * FROM Warehouse WHERE location = '" + splitdata[7] + "'", splitdata[1]);
                contract_code = await execquery("SELECT * FROM SKD.Product WHERE ProductCode = '" + splitdata[9] + "'", splitdata[1]);
                tradeaccount = await execquery("SELECT * FROM SKD.Investor WHERE Code = '" + splitdata[10] + "'", splitdata[1]);
                if (depositor != "ok" || warehouse != "ok" || contract_code != "ok" || tradeaccount != "ok")
                {
                    break;
                }
            }
            if (depositor == "ok" && warehouse == "ok" && contract_code == "ok" && tradeaccount == "ok")
            {
                return ("Depositor ✅\nWarehouse ✅\nContract Code ✅\nTrade Account ✅\nQuantity ✅\nSilahkan upload data di tinmarket.id");
            }
            else
            {
                return ("Depositor : "+depositor+"\n" +
                        "Warehouse : "+warehouse+"\n"+
                        "Contract Code : "+ contract_code +"\n"+
                        "Trade Account : "+tradeaccount+"\n");
            }
        }
        private static string SendMessage(string chatId, string body)
        {
            var client = new RestClient("https://api.1msg.io/127354/sendMessage?token=jkdjtwjkwq2gfkac");

            RestRequest requestWa = new RestRequest("https://api.1msg.io/127354/sendMessage?token=jkdjtwjkwq2gfkac", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("chatId", chatId);
            requestWa.AddParameter("body", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);

        }
        private static void sendTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;

            var client = new RestClient("https://api.telegram.org/bot5278461864:AAGwGEV3aJx8ZFKFfRaFB6TMc4mOem-uoc8/sendMessage?chat_id=" + chatId + "&text=" + body);
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot5278461864:AAGwGEV3aJx8ZFKFfRaFB6TMc4mOem-uoc8/sendMessage?chat_id=" + chatId + "&text=" + body, Method.Get);
            requestWa.Timeout = -1;
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);
        }
        public static async Task<string> execquery(string query, string index)
        {
            string confirm = "ok";
            var data = new Data();
            using (SqlConnection connection = new SqlConnection(data._connectionString))
            {
                connection.Open();

                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = query;

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows == false)
                        {
                            confirm = ("Data not match, Please check data line "+ index);
                        }
                    }
                }
                connection.Close();
            }
            return (confirm);
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