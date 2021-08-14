using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Google.Apis.Gmail.v1.Data;

namespace GmailAPI.APIHelper
{
    public static class GmailAPIHelper
    {
        static string[] Scores = { GmailService.Scope.MailGoogleCom };
        static string ApplicationName = "Gmail API Application";

        public static GmailService GetService()
        {
            UserCredential credential;
            using (FileStream stream = new FileStream(Convert.ToString(ConfigurationManager.AppSettings["ClientInfo"]),
                FileMode.Open, FileAccess.Read))
            {
                string FolderPath = Convert.ToString(ConfigurationManager.AppSettings["CredentialsInfo"]);
                string FilePath = Path.Combine(FolderPath, "APITokenCredentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scores,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(FilePath, true)).Result;
            }

            //Create Gmail API service
            GmailService service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });
            return service;
        }

        public static string MsgNestedParts(IList<MessagePart> payloadParts)
        {
            string str = string.Empty;
            if (payloadParts.Count() < 0)
            {
                return string.Empty;
            }
            else
            {
                IList<MessagePart> plainTestMail = payloadParts.Where(parts => parts.MimeType == "text/plain").ToList();
                IList<MessagePart> attachmentMail = payloadParts.Where(parts => parts.MimeType == "multipart/alternative").ToList();

                if (plainTestMail.Count() > 0)
                {
                    foreach (MessagePart eachPart in plainTestMail)
                    {
                        if (eachPart.Parts == null)
                        {
                            if (eachPart.Body != null && eachPart.Body.Data != null)
                            {
                                str += eachPart.Body.Data;
                            }
                        }
                        else
                        {
                            return MsgNestedParts(eachPart.Parts);
                        }
                    }
                }

                if (attachmentMail.Count() > 0)
                {
                    foreach (MessagePart eachPart in attachmentMail)
                    {
                        if (eachPart.Parts == null)
                        {
                            if (eachPart.Body != null && eachPart.Body.Data != null)
                            {
                                str += eachPart.Body.Data;
                            }
                        }
                        else
                        {
                            return MsgNestedParts(eachPart.Parts);
                        }
                    }
                }
                return str;
            }
        }

        public static List<string> GetAttachments(string userId, string messageId, string outputDir)
        {
            try
            {
                List<string> FileName = new List<string>();
                GmailService gmailService = GetService();
                Message message = gmailService.Users.Messages.Get(userId, messageId).Execute();
                IList<MessagePart> messageParts = message.Payload.Parts;

                foreach (var messagePart in messageParts)
                {
                    if (!String.IsNullOrEmpty(messagePart.Filename))
                    {
                        string attachId = messagePart.Body.AttachmentId;
                        MessagePartBody attachPart =
                            gmailService.Users.Messages.Attachments.Get(userId, messageId, attachId).Execute();
                        byte[] data = Base64ToByte(attachPart.Data);
                        File.WriteAllBytes(Path.Combine(outputDir, messagePart.Filename), data);
                    }
                }

                return FileName;
            }
            catch (Exception exception)
            {
                Console.WriteLine("An error occurred: " + exception.Message);
                return null;
            }
        }


        public static string Base64Decode(string Base64Test)
        {
            string EncodTxt = string.Empty;
            //STEP-1: Replace all special Character of Base64Test
            EncodTxt = Base64Test.Replace("-", "+");
            EncodTxt = EncodTxt.Replace("_", "/");
            EncodTxt = EncodTxt.Replace(" ", "+");
            EncodTxt = EncodTxt.Replace("=", "+");

            //STEP-2: Fixed invalid length of Base64Test
            if (EncodTxt.Length % 4 > 0) { EncodTxt += new string('=', 4 - EncodTxt.Length % 4); }
            else if (EncodTxt.Length % 4 == 0)
            {
                EncodTxt = EncodTxt.Substring(0, EncodTxt.Length - 1);
                if (EncodTxt.Length % 4 > 0) { EncodTxt += new string('+', 4 - EncodTxt.Length % 4); }
            }

            //STEP-3: Convert to Byte array
            byte[] ByteArray = Convert.FromBase64String(EncodTxt);

            //STEP-4: Encoding to UTF8 Format
            return Encoding.UTF8.GetString(ByteArray);
        }

        public static byte[] Base64ToByte(string Base64Test)
        {
            string EncodTxt = string.Empty;
            //STEP-1: Replace all special Character of Base64Test
            EncodTxt = Base64Test.Replace("-", "+");
            EncodTxt = EncodTxt.Replace("_", "/");
            EncodTxt = EncodTxt.Replace(" ", "+");
            EncodTxt = EncodTxt.Replace("=", "+");

            //STEP-2: Fixed invalid length of Base64Test
            if (EncodTxt.Length % 4 > 0) { EncodTxt += new string('=', 4 - EncodTxt.Length % 4); }
            else if (EncodTxt.Length % 4 == 0)
            {
                EncodTxt = EncodTxt.Substring(0, EncodTxt.Length - 1);
                if (EncodTxt.Length % 4 > 0) { EncodTxt += new string('+', 4 - EncodTxt.Length % 4); }
            }

            //STEP-3: Convert to Byte array
            return Convert.FromBase64String(EncodTxt);
        }

     

        public static void MsgMarkAsRead(string HostEmailAdress, string MsgId)
        {
            //Message marks as read after reading message
            ModifyMessageRequest modifyMessageRequest = new ModifyMessageRequest();
            modifyMessageRequest.AddLabelIds = null;
            modifyMessageRequest.RemoveLabelIds = new List<string> { "UNREAD" };
            GetService().Users.Messages.Modify(modifyMessageRequest, HostEmailAdress, MsgId).Execute();
        }

        
    }
}
