using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;


namespace OLEemailRetrieve
{
    internal class Program
    {
        static void Main()
        {
            var ssConversion = new Conversion();
            
            var dt = DateTime.Now.ToString("MMddyyyyhhmmss");
            var path = @"ImportedFiles\";
            var log = @"Logs\";
            var logfilename = dt + ".txt";

            try
            {
                if (!Directory.Exists(log))
                {
                    Directory.CreateDirectory(log);
                }
            }

            catch (Exception)
            {
                // ignored
            }

            var filestream = new FileStream(log + logfilename, FileMode.Create);
            var streamwriter = new StreamWriter(filestream) {AutoFlush = true};
            Console.SetOut(streamwriter);
            Console.SetError(streamwriter);
            
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                streamwriter.WriteLine("ImportedFiles directory created!");
            }
            catch (Exception ex)
            {
                streamwriter.WriteLine(ex.ToString());
            }

            // Email App Settings
            var mEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            var mEmailSubject = ConfigurationManager.AppSettings["EmailSubject"];
            
            var service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            if (ConfigurationManager.AppSettings["UseDefault"] != "true")
            {
                service.Credentials = new WebCredentials(ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Password"], ConfigurationManager.AppSettings["Domain"]);
                service.TraceEnabled = true;
                service.TraceFlags = TraceFlags.All;

                try
                {
                    service.AutodiscoverUrl(ConfigurationManager.AppSettings["Address"], RedirectionUrlValidationCallback);
                }
                catch
                {
                    streamwriter.WriteLine("Not a valid [Address] in app config");
                }
                
            }
            else
            {
                service.UseDefaultCredentials = true;
                try
                {
                    service.AutodiscoverUrl(mEmailFrom, RedirectionUrlValidationCallback);
                }
                catch
                {
                    streamwriter.WriteLine("Not a valid [EMAILFROM] in app config");
                }
                
            }

            // Add a search filter that searches on the body or subject.
            var searchFilterCollection = new List<SearchFilter>
            {
                new SearchFilter.ContainsSubstring(ItemSchema.Subject, mEmailSubject)
            };
            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchFilterCollection.ToArray());

            // Create a view with a page size of 10.
            var view = new ItemView(10) {PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject)};

            // Identify the Subject and DateTimeReceived properties to return.
            // Indicate that the base property will be the item identifier

            // Order the search results by the DateTimeReceived in descending order.
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
            
            view.Traversal = ItemTraversal.Shallow;

            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

            foreach (var message in findResults.Items.Select(item => EmailMessage.Bind(service, new ItemId(item.Id.ToString()), new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments))))
            {
                // Iterate through the attachments collection and load each attachment.
                if (message.Attachments.Count == 1)
                {
                    foreach (var fileAttachment in message.Attachments.OfType<FileAttachment>())
                    {
                        fileAttachment.Load();
                        streamwriter.WriteLine("Attachment name: " + fileAttachment.Name);

                        // Load attachment contents into a file.
                        fileAttachment.Load(path + fileAttachment.Name);

                        var fileName = path + fileAttachment.Name;
                        ssConversion.ConvertXls(fileName);
                    }
                }
                else
                {
                    foreach (FileAttachment fileAttachment in from attachment in message.Attachments where attachment is FileAttachment && attachment.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" select attachment as FileAttachment)
                    {
                        // Load the file attachment into memory and print out its file name.
                        fileAttachment.Load();
                        streamwriter.WriteLine("Attachment name: " + fileAttachment.Name);

                        // Load attachment contents into a file.
                        fileAttachment.Load(path + fileAttachment.Name);

                        var fileName = path + fileAttachment.Name;
                        ssConversion.ConvertXls(fileName);
                    }
                }
            }    
        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
