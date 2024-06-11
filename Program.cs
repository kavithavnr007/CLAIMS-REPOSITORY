using System;
using System.Security;
using Microsoft.SharePoint.Client;
using System.IO;
using File = System.IO.File;
using System.Configuration;

namespace AttachmentMapping
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = ConfigurationManager.AppSettings["SiteUrl"];
            string listName = ConfigurationManager.AppSettings["ListName"];
            //string siteUrl = siteUrlHR;
            //string listName = listName;   
            int detailLineValue = 1;

            string username = "Connectadmin@sony.onmicrosoft.com";
            string password = "THX@v0lum3";

            SecureString securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }

            string localFolderPath = @ConfigurationManager.AppSettings["FilePath"];
            
            // Connect to SharePoint site
            using (ClientContext context = new ClientContext(siteUrl))
            {
                // Provide credentials
                context.Credentials = new SharePointOnlineCredentials(username, securePassword);

                // Get the list
                List list = context.Web.Lists.GetByTitle(listName);

                // Define CAML query to filter items where detail_line equals 1
                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='detail_line'/><Value Type='Text'>" + detailLineValue + "</Value></Eq></Where></Query></View>";

                // Set query row limit to 5000
                query.ListItemCollectionPosition = null;
                query.ViewXml = "<View Scope='RecursiveAll'><Query>" + query.ViewXml + "</Query><RowLimit Paged='TRUE'>5000</RowLimit></View>";

                do
                {
                    ListItemCollection items = list.GetItems(query);
                    context.Load(items);
                    context.ExecuteQuery();

                    foreach (ListItem item in items)
                    {
                        // Get claim_id from the current item
                        string claimId = item["claim_id"].ToString();

                        // Check if the folder exists
                        if (Directory.Exists(localFolderPath))
                        {
                            // Get files in the folder
                            string[] files = Directory.GetFiles(localFolderPath);

                            // Loop through each file in the folder
                            foreach (string filePath in files)
                            {
                                // Check if the file name starts with the desired claim ID
                                if (Path.GetFileName(filePath).StartsWith(claimId))
                                {
                                    Console.WriteLine($"Found document '{Path.GetFileName(filePath)}' in the local folder.");

                                    // Check if the attachment with the same name already exists
                                    if (!AttachmentExists(context, item, Path.GetFileName(filePath)))
                                    {
                                        // Read the file content from local drive
                                        byte[] fileContent = File.ReadAllBytes(filePath);

                                        // Add attachment to list item
                                        AttachmentCreationInformation attachmentInfo = new AttachmentCreationInformation
                                        {
                                            FileName = Path.GetFileName(filePath),
                                            ContentStream = new MemoryStream(fileContent)
                                        };

                                        Attachment attachment = item.AttachmentFiles.Add(attachmentInfo);
                                        context.ExecuteQuery();

                                        Console.WriteLine("Attachment added successfully.");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Attachment '{Path.GetFileName(filePath)}' already exists. Skipping...");
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Local folder does not exist.");
                        }
                    }

                    query.ListItemCollectionPosition = items.ListItemCollectionPosition;
                }
                while (query.ListItemCollectionPosition != null);

                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }

        // Check if the attachment with the specified name already exists for the list item
        static bool AttachmentExists(ClientContext context, ListItem listItem, string attachmentName)
        {
            context.Load(listItem.AttachmentFiles);
            context.ExecuteQuery();

            foreach (var attachment in listItem.AttachmentFiles)
            {
                if (attachment.FileName == attachmentName)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
