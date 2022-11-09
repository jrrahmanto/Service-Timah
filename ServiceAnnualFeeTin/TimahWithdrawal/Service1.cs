using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Newtonsoft.Json;
using RestSharp;
using static ConsoleApplication2.ResponsWhatsapp;
using System.Net;
using System.Globalization;
using System.Text.RegularExpressions;

namespace TimahWithdrawal
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            Timer timer = new Timer();
            //timer.Interval = 86400000; // 1 hari
            timer.Interval = 15000;  //15 detik
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
        }
        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            bool format_old = false;
            List<int> messageNumber = new List<int>();

            string chat_id = "6289630870658-1628583102@g.us";
            int lognumber = 0;
            string file = AppDomain.CurrentDomain.BaseDirectory + "\\logLastMessage.txt";
            string[] text = File.ReadAllLines(file);
            messageNumber.Add(Convert.ToInt32(text[0]));
            string message = GetMessageList(chat_id, messageNumber[messageNumber.Count - 1]);
            ResponseChat rc = JsonConvert.DeserializeObject<ResponseChat>(message);

            foreach (MessageChat mc in rc.messages)
            {
                lognumber = mc.messageNumber;
                //masukin last message numbernya
                string text2 = System.IO.File.ReadAllText(file);
                text2 = lognumber.ToString();
                System.IO.File.WriteAllText(file, text2);
                SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog : Message Number"+ lognumber);

                if (mc.caption != null)
                {
                    if ((mc.caption.Contains("1220010133075")) && (mc.body.EndsWith(".csv")))
                    {
                        try
                        {
                            SendMessage(chat_id, "Processing upload proof of payment start " + DateTime.Now.ToString("HH:mm:ss"));
                            SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog : Processing upload proof of payment start");

                            List<data_csv> dataCsv = new List<data_csv>();
                            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Report\\wd.xlsx";
                            try
                            {
                                using (WebClient wc = new WebClient())
                                {
                                    wc.DownloadFile(new System.Uri(mc.body), path);
                                }
                            }
                            catch (Exception ex)
                            {

                            }
                            string[] arrBuktiTf = File.ReadAllLines(path);
                            int j = 0;
                            if (arrBuktiTf[0].Contains(";"))
                            {
                                for (int i = 1; i < arrBuktiTf.Length; i++)
                                {
                                    string[] arrDataTf = arrBuktiTf[i].Split(';');
                                    dataCsv.Add(new data_csv
                                    {
                                        AccountNo = (arrDataTf[0].ToString()),
                                        Date = (arrDataTf[2].ToString().Remove(arrDataTf[2].Length - 9)),
                                        ValDate = (arrDataTf[2].ToString().Remove(arrDataTf[2].Length - 9)),
                                        Description1 = (arrDataTf[3].ToString()),
                                        Description2 = (arrDataTf[4].ToString()),
                                        Debit = (arrDataTf[6].ToString()),
                                        Credit = (arrDataTf[5].ToString()),
                                    });
                                }
                            }
                            else
                            {
                                format_old = true;
                                StreamReader sr = new StreamReader(path);
                                while (!sr.EndOfStream)
                                {
                                    MatchCollection matches = new Regex("((?<=\")[^\"]*(?=\"(,|$)+)|(?<=,|^)[^,\"]*(?=,|$))").Matches(sr.ReadLine());
                                    string[] rows = new string[matches.Count + 1];
                                    int i = 1;
                                    foreach (var item in matches)
                                    {
                                        rows[i] = item.ToString();
                                        i++;
                                    }

                                    if (j >= 1)
                                    {
                                        dataCsv.Add(new data_csv
                                        {
                                            AccountNo = (rows[1].ToString()),
                                            Date = (rows[2].ToString()),
                                            ValDate = (rows[3].ToString()),
                                            TransactionCode = (rows[4].ToString()),
                                            Description1 = (rows[5].ToString()),
                                            Description2 = (rows[6].ToString()),
                                            ReferenceNo = (rows[7].ToString()),
                                            Debit = (rows[8].ToString()),
                                            Credit = (rows[9].ToString()),
                                        });
                                    }
                                    j++;
                                }
                            }

                            var dt = new ConsoleApplication2.BankTrasferTableAdapters.Bank_TransferTableAdapter();

                            foreach (var item in dataCsv)
                            {
                                SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog : Before parsing" + item.Date + item.Description1 + item.Description2 + item.Credit);
                              
                                if (item.Credit != ".00")
                                {
                                    Decimal number_credit = Convert.ToDecimal(item.Credit.Replace(",", string.Empty).Replace(".",","));
                                    DateTime date = new DateTime();
                                    var date1 = DateTime.Now.ToString("dd MMMM yyyy");
                                    if (format_old == true)
                                    {
                                        date = DateTime.ParseExact(item.Date, "dd/MM/yy", CultureInfo.InvariantCulture);
                                    }
                                    else
                                    {
                                        date = DateTime.ParseExact(item.Date, "dd MMMM yyyy", CultureInfo.InvariantCulture);
                                    }
                                    var dr = dt.GetDataByFilter(date.ToString("yyyy-MMMM-dd"), item.Description2, number_credit);

                                    SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog After Parsing: " + date + item.Description1 + item.Description2 + number_credit);
                                    WriteToFile(date + item.Description1 + item.Description2 + number_credit);
                                    if (dr.Count == 0)
                                    {
                                        int newId = -1;
                                        Object retObj = null;
                                        retObj = dt.InsertQuery(date.ToString("yyyy-MM-dd"), item.Description1, item.Description2, number_credit, DateTime.Now, "Robot KBI");
                                        newId = Convert.ToInt32(retObj);

                                        SendMessage(chat_id, newId + "#withdrawal#\nThis Amount Transfer : " + number_credit + "\nPlease answer this message with the format :\n1.ExchangeReff\n2.Amount\n3.Use Secfund(Yes/No)\n\nTimeStamp: " + DateTime.Now.ToString("HH:mm:ss"));
                                        SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog After Parsing: Success Insert ID Transfer = "+newId);

                                    }
                                    else
                                    {
                                        WriteToFile("data udah ada");
                                        //SendMessage(chat_id, "This data has been uploaded " + DateTime.Now.ToString("HH:mm:ss"));
                                    }
                                }
                            }
                            SendMessage(chat_id, "Processing upload proof of payment success " + DateTime.Now.ToString("HH:mm:ss"));
                        }
                        //}
                        catch (Exception ex)
                        {
                            WriteToFile(ex.Message);
                            SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog Process upload proof of payment failed " + ex.Message);

                            SendMessage(chat_id, "Process upload proof of payment failed " + ex.Message + " " + DateTime.Now.ToString("HH:mm:ss"));
                        }
                    }

                }
            }
        }
        protected override void OnStop()
        {
        }
        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        private static string GetMessageList(string chatId, int lastMessageNumber)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/messages?token=jkdjtwjkwq2gfkac&lastMessageNumber=" + lastMessageNumber + "&chatId=" + chatId);
            client.Timeout = -1;
            var requestWa = new RestRequest(Method.GET);
            IRestResponse responseWa = client.Execute(requestWa);
            return responseWa.Content;
        }

        private static void SendMessage(string chatId, string body)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac");
            client.Timeout = -1;
            var requestWa = new RestRequest(Method.POST);
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("chatId", chatId);
            requestWa.AddParameter("body", body);
            IRestResponse responseWa = client.Execute(requestWa);
            Console.WriteLine(responseWa.Content);
        }
        private static string SendTelegram(string chatId, string message)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12;

            var client = new RestClient("https://api.telegram.org/bot2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U/sendMessage?chat_id=" + chatId + "&text=" + message);
            client.Timeout = -1;
            var requestWa = new RestRequest(Method.GET);
            IRestResponse responseWa = client.Execute(requestWa);
            return responseWa.Content;
        }
    }
}
