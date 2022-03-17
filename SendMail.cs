using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EveryTeacher
{
    class SendMail
    {
        public static string CFG_HEADER_SENDNAME = "收件人名稱";
        public static string CFG_HEADER_SENDTO = "收件人信箱";
        public static string CFG_HEADER_CC = "副本";
        public static string CFG_HEADER_ATTACH = "附件檔案";
        public static string CFG_HEADER_SUBJECT = "信件主旨";
        public static string CFG_HEADER_BODY = "信件內容";

        public string SendName = "";
        public string Sendto = "";
        public string CC = "";
        public string Attach = "";
        public string Title = "";
        public string Subject = "";

        public SendMail()
        {
            this.SendName = "";
            this.Sendto = "";
            this.CC = "";
            this.Attach = "";
            this.Title = "";
            this.Subject = "";
        }

        public static string[] getAllHeadersText()
        {
            string[] arr = {CFG_HEADER_SENDNAME , CFG_HEADER_SENDTO,CFG_HEADER_CC,
                CFG_HEADER_ATTACH, CFG_HEADER_SUBJECT, CFG_HEADER_BODY };
            return arr;
        }
        public string[] getAllItemsText()
        {
            string[] arr = {this.SendName, this.Sendto, this.CC,
                this.Attach, this.Title, this.Subject };
            return arr;
        }
    }
}
