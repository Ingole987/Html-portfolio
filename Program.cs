using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace OAuth2
{
    class Alternative8
    {
        //const string TenantId = "72c490b4-b8a8-4d0c-9915-f6b23e15b7f5";
        //const string AppId = "f3b4e99f-6227-4121-af04-b82d1ed35430";
        //const string AppSecret = "iDT8Q~G2TOdcYPaHSfM~Ga0W0FgQGGaYVZ3SqaZZ";
        //const string Username = "iconnect.test@provana.com";
        //const string Password = "Supp0rt@12301";
        static async Task Main(string[] args)
        {
            Console.Write("Program Started");
            await RetrieveEmail();

            Console.Write("Press ENTER to end this program");
            Console.ReadLine();
        }

        static async Task RetrieveEmail()
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var tenantId = "72c490b4-b8a8-4d0c-9915-f6b23e15b7f5";
            var clientId = "f3b4e99f-6227-4121-af04-b82d1ed35430";
            var clientSecret = "iDT8Q~G2TOdcYPaHSfM~Ga0W0FgQGGaYVZ3SqaZZ";
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret);

            string after = "2022-8-1";
            DateTime oDate = DateTime.Parse(after);
            var after1 = new DateTime(oDate.Year, oDate.Month, oDate.Day).ToString("yyyy-MM-dd");
            var filter = $"(receivedDateTime gt {after1})";


            //var before = new DateTime(oDate.Year, oDate.Month, oDate.Day).AddDays(1).ToString("yyyy-MM-dd");
            //var filter = $"(receivedDateTime gt {after1}) and (receivedDateTime le {before}) ";

            GraphServiceClient graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            var inboxMessages = await graphClient
                                    .Users["iconnect.test@provana.com"]
                                    //.Me
                                    .MailFolders["inbox"]
                                    .Messages
                                    .Request()
                                    .Filter(filter)
                                    //.Top(200)
                                    .Expand("attachments")
                                    .GetAsync();


            var totalListItems = new List<Message>();
            totalListItems.AddRange(inboxMessages.CurrentPage);


            while (inboxMessages.NextPageRequest != null)
            {
                inboxMessages = inboxMessages.NextPageRequest.GetAsync().Result;
                totalListItems.AddRange(inboxMessages.CurrentPage);
            }

            if (totalListItems.Count > 0)
            {
                foreach (var item in inboxMessages.CurrentPage)
                {
                    item.Body.Content.ToString();

                    item.Subject.ToString();

                    foreach (var At in item.Attachments)
                    {
                        var At1 = (FileAttachment)At;
                        if (!String.IsNullOrEmpty(At1.ContentId))
                        {
                            At1.ContentId.ToString();
                            At1.ContentType.ToString();
                        }
                        var downloadPath = @"C:\IC247\SourceCode\OAuth2\OAuth2\bin\Debug\Attachment";
                        var fileName = DateTime.Now.ToFileTime() + "_" + At1.Name;
                        System.IO.File.WriteAllBytes(System.IO.Path.Combine(downloadPath, fileName), At1.ContentBytes);
                    }


                    foreach (var BCC in item.BccRecipients)
                    {
                        string strBCC1 = BCC.EmailAddress.Address;
                        string strBCC = BCC.EmailAddress.Name;
                    }

                    foreach (var CC in item.CcRecipients)
                    {
                        string strCC = CC.EmailAddress.Address;
                        string strBCC = CC.EmailAddress.Name;
                    }

                    item.BodyPreview.ToString();

                    item.From.EmailAddress.Address.ToString();
                    item.From.EmailAddress.Name.ToString();


                    item.ReceivedDateTime.Value.UtcDateTime.ToString();
                    item.ReceivedDateTime.Value.Date.ToString();
                    item.ReceivedDateTime.Value.LocalDateTime.ToString();


                    item.ReplyTo.ToString();

                    item.Sender.EmailAddress.Address.ToString();
                    item.Sender.EmailAddress.Name.ToString();

                    item.SentDateTime.ToString();

                    item.ToRecipients.ToString();
                }

            }

            //}
            //return "Success";
        }
    }
}


