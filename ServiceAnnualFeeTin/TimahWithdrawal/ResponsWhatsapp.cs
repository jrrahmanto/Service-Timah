using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    class ResponsWhatsapp
    {
        public class ResponseChat
        {
            public IEnumerable<MessageChat> messages { get; set; }
            public int lastMessageNumber { get; set; }
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
        public class data_csv
        {
            public string AccountNo { get; set; }
            public string Date { get; set; }
            public string ValDate { get; set; }
            public string TransactionCode { get; set; }
            public string Description1 { get; set; }
            public string Description2 { get; set; }
            public string ReferenceNo { get; set; }
            public string Debit { get; set; }
            public string Credit { get; set; }

        }
    }
}
