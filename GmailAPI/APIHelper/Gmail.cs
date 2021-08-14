﻿using System;
using System.Collections.Generic;

namespace GmailAPI.APIHelper
{
    public class Gmail
    {
        public string From { get; set; }

        public string Body { get; set; }

        public DateTime MailDateTime { get; set; }

        public List<string> Attachments { get; set; }
        public string MsgID { get; set; }

    }
}