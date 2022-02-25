using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EveryTeacher
{
    class SendMail
    {
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
    }
}
