using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EveryTeacher
{
    class SendMail
    {
        public string SendName { get; set; }
        public string Sendto { get; set; }
        public string CC { get; set; }
        public string Attach { get; set; }
        public string Title { get; set; }
        public string Subject { get; set; }

        public SendMail()
        {
            this.SendName = "";
            this.Sendto = "";
            this.CC = "";
            this.Attach = "";
            this.Title = "";
            this.Subject = "";
        }
    }
}
