using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Channels;
using GmailAPI.APIHelper;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;


namespace GmailAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                List<Gmail> MailList = GetAllEmails(Convert.ToString(
                    ConfigurationManager.AppSettings["HostAddress"]));
                ;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Error: " + exception);
            }
        }

        public static List<Gmail> GetAllEmails(string HostEmailAddress)
        {
            try
            {
                GmailService gmailService = GmailAPIHelper.GetService();
                List<Gmail> emailList = new List<Gmail>();
                UsersResource.MessagesResource.ListRequest listRequest =
                    gmailService.Users.Messages.List(HostEmailAddress);
                listRequest.LabelIds = "INBOX";
                listRequest.IncludeSpamTrash = false;
                listRequest.Q = "is:unread"; //only for unread e-mails

                //Get All Emails
                ListMessagesResponse listMessagesResponse = listRequest.Execute();

                if (listMessagesResponse != null && listMessagesResponse.Messages != null)
                {
                    //Loop through each email and get what fields I want
                    foreach (Message msg in listMessagesResponse.Messages)
                    {
                        //Message marks as read after reading message
                        GmailAPIHelper.MsgMarkAsRead(HostEmailAddress, msg.Id);

                        UsersResource.MessagesResource.GetRequest Message =
                            gmailService.Users.Messages.Get(HostEmailAddress, msg.Id);
                        Console.WriteLine("\n --New Mail -- ");
                        Console.WriteLine("Step-1: Message ID: " + msg.Id);

                        //Make another request for that email id
                        Message msgContent = Message.Execute();

                        if (msgContent != null)
                        {
                            string FromAddress = string.Empty;
                            string Date = string.Empty;
                            string Subject = string.Empty;
                            string MailBody = string.Empty;
                            string ReadableText = string.Empty;

                            //Loop through the headers and get fields we need (Subject, Mail)
                            foreach (var messagePart in msgContent.Payload.Headers)
                            {
                                if (messagePart.Name == "From")
                                {
                                    FromAddress = messagePart.Value;
                                }
                                else if (messagePart.Name == "Date")
                                {
                                    Date = messagePart.Value;
                                }
                                else if (messagePart.Name == "Subject")
                                {
                                    Subject = messagePart.Value;
                                }
                            }

                            //Read mail body
                            Console.WriteLine("Step-2: Read Mail Body");
                            List<string> FileName = GmailAPIHelper.GetAttachments(
                                HostEmailAddress,
                                msg.Id,
                                Convert.ToString(ConfigurationManager.AppSettings["GmailAttach"]));

                            if (FileName.Count() > 0)
                            {
                                foreach (var eachFile in FileName)
                                {
                                    //Get user Id using from Email Address
                                    string[] rectifyFromAddress = FromAddress.Split(' ');
                                    string fromAdd = rectifyFromAddress[rectifyFromAddress.Length - 1];

                                    if (!String.IsNullOrEmpty(fromAdd))
                                    {
                                        fromAdd = fromAdd.Replace("<", string.Empty);
                                        fromAdd = fromAdd.Replace(">", string.Empty);
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Step-3: Mail has no attachments.");
                            }

                            //Read mail body
                            MailBody = String.Empty;
                            if (msgContent.Payload.Parts == null && msgContent.Payload.Parts != null)
                            {
                                MailBody = msgContent.Payload.Body.Data;
                            }
                            else
                            {
                                MailBody = GmailAPIHelper.MsgNestedParts(msgContent.Payload.Parts);
                            }

                            //BASE64 TO READABLE TEXT--------------------------------------------------------------------------------
                            ReadableText = string.Empty;
                            ReadableText = GmailAPIHelper.Base64Decode(MailBody);

                            Console.WriteLine("STEP-4: Identifying & Configure Mails.");

                            if (!string.IsNullOrEmpty(ReadableText))
                            {
                                Gmail GMail = new Gmail();
                                GMail.From = FromAddress;
                                GMail.Body = ReadableText;
                                GMail.MailDateTime = Convert.ToDateTime(Date);
                                emailList.Add(GMail);
                            }
                        }
                    }
                }
                return emailList;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);

                return null;
            }
        }
    }
}
