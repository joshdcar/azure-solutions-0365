using System;
using System.Linq;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.WindowsAzure.Storage;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.Configuration;
using Microsoft.WindowsAzure.Storage.Queue;
using SmartPhotoLibrary.Azure.Functions.Models;
using Newtonsoft.Json;
using Microsoft.WindowsAzure.Storage.Table;

namespace SmartPhotoLibrary.Azure.Functions
{
    public static class PhotoNotifications
    {
        private static TraceWriter _log = null;

        [FunctionName("PhotoNotifications")]        
        public static async Task Run([QueueTrigger("photonotifications", Connection = "AppStorage")]string requestMessage, TraceWriter log)
        {
            _log = log;

            _log.Info($"Queue event for photonotification. Message: {requestMessage}");

            var request = JsonConvert.DeserializeObject<NotificationModel>(requestMessage);

            await GetChanges(request);

        }

        /// <summary>
        /// Get the lastest changes to the list based on the last change token time
        /// </summary>
        /// <param name="notification"></param>
        /// <returns></returns>
        private static async Task GetChanges(NotificationModel notification)
        {
            // *** WARNING ****
            //Using a username\password for site authentication is not considered a best practice
            //and is difficult to secure and can often end up in source (bad)
            //Strongly Consider using a Azure AD Application authentication

            string siteUrl = ConfigurationManager.AppSettings["spurl"];
            string userName = ConfigurationManager.AppSettings["spusername"];
            string password = ConfigurationManager.AppSettings["sppassword"];

            OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();

            using (var clientContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, userName, password))
            {
                ListCollection lists = clientContext.Web.Lists;
                Guid listId = new Guid(notification.Resource);

                IEnumerable<List> results = clientContext.LoadQuery<List>(lists.Where(lst => lst.Id == listId));
                clientContext.ExecuteQueryRetry();

                List changeList = results.FirstOrDefault();

                if (changeList == null)
                {
                    return;
                }

                // grab last change token from storage
                var changeHistory = GetWebhookHistory(changeList.Id);
                var lastChangeToken = string.Empty;

                //Assign our change token 
                if (changeHistory != null)
                {
                    lastChangeToken = changeHistory.LastChangeToken;
                }
                else
                {
                    lastChangeToken = string.Format("1;3;{0};{1};-1", notification.Resource, DateTime.Now.AddMinutes(-15).ToUniversalTime().Ticks.ToString());
                }

                _log.Info($"Last change request token: {lastChangeToken}");

                ChangeQuery changeQuery = new ChangeQuery(false, true);
                changeQuery.Item = true;
                changeQuery.ChangeTokenStart = new ChangeToken() { StringValue = lastChangeToken };


                var changes = changeList.GetChanges(changeQuery);
                clientContext.Load(changes);
                clientContext.ExecuteQueryRetry();

                foreach (Change change in changes)
                {
                    //We only want to responde to Add\Update Events
                    if (change.ChangeType == ChangeType.Add || change.ChangeType == ChangeType.Update)
                    {
                        //We don't want to re-analyze existing photos so we're going to check
                        //Also without some sort of facility like this we can cause an endless update\analyze\update loop
                        if (!AlreadyAnalyzed(change as ChangeItem))
                        {
                            //Queue up our Photo Processing Request
                            await QueuePhotoProcessing(changeList, change);
                        }

                    }
                }

                //Save our change Token to Storage
                changeHistory = new WebhookHistoryModel()
                {
                    Id = new Guid(notification.SubscriptionId),
                    ListId = listId,
                    LastChangeToken = string.Format("1;3;{0};{1};-1", notification.Resource, DateTime.Now.ToUniversalTime().Ticks.ToString())
                };

                SaveWebhookHistory(changeHistory);

            }

        }

        /// <summary>
        /// Check and see if the photo has already been analyzed before.
        /// </summary>
        /// <param name="changeItem"></param>
        /// <returns></returns>
        private static bool AlreadyAnalyzed(ChangeItem changeItem)
        {
            var analyzed = true; //we're going to prevent by default

            // *** WARNING ****
            //Using a username\password for site authentication is not considered a best practice
            //and is difficult to secure and can often end up in source (bad)
            //Strongly Consider using a Azure AD Application authentication

            string siteUrl = ConfigurationManager.AppSettings["spurl"];
            string userName = ConfigurationManager.AppSettings["spusername"];
            string password = ConfigurationManager.AppSettings["sppassword"];

            OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();

            using (var clientContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, userName, password))
            {
                ListCollection lists = clientContext.Web.Lists;

                var listID = changeItem.ListId;

                IEnumerable<List> listResults = clientContext.LoadQuery<List>(lists.Where(lst => lst.Id == listID));
                clientContext.ExecuteQueryRetry();

                var photoLibrary = listResults.FirstOrDefault();

                if (photoLibrary != null)
                {
                    var query = new CamlQuery();
                    query.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>{changeItem.ItemId}</Value></Eq></Where></Query><ViewFields><FieldRef Name='Analyzed' /><FieldRef Name='Analyzed' /></ViewFields></View>";

                    var items = photoLibrary.GetItems(query);

                    clientContext.Load(items, includes => includes.Include(i => i["Analyzed"]));
                    clientContext.ExecuteQuery();

                    var item = items.FirstOrDefault();

                    if (item != null)
                    {
                        if (item["Analyzed"] != null)
                        {
                            analyzed = bool.Parse(item["Analyzed"].ToString());
                        }
                    }

                }
            }

            return analyzed;
        }

        /// <summary>
        /// Submit the photo analysis request to the queue
        /// </summary>
        /// <param name="changeList"></param>
        /// <param name="change"></param>
        /// <returns></returns>
        private static async Task QueuePhotoProcessing(List changeList, Change change)
        {

            string storageAccountConnectionString = ConfigurationManager.AppSettings["AppStorage"];
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(storageAccountConnectionString);

            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference("photorequests");

            queue.CreateIfNotExists();

            var request = new PhotoRequestModel() { Id = (change as ChangeItem).ItemId, ListId = changeList.Id };

            string message = JsonConvert.SerializeObject(request);

            _log.Info($"Queueing Photo Request. Message: {message}");

            await queue.AddMessageAsync(new CloudQueueMessage(message));

        }

        /// <summary>
        /// Retrieve webhook processing history from Azure Table Storage
        /// This is important so we don't reprocess events over again
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        private static WebhookHistoryModel GetWebhookHistory(Guid id)
        {
            string storageAccountConnectionString = ConfigurationManager.AppSettings["AppStorage"];
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(storageAccountConnectionString);

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
            CloudTable webhookTable = tableClient.GetTableReference("webhookhistory");

            webhookTable.CreateIfNotExists();

            TableOperation retrieveOperation = TableOperation.Retrieve<WebhookHistoryModel>("ListHistory", id.ToString());

            // Execute the retrieve operation.
            TableResult retrievedResult = webhookTable.Execute(retrieveOperation);

            return retrievedResult.Result as WebhookHistoryModel;
        }

        /// <summary>
        /// Save the latest history to azure table storage
        /// </summary>
        /// <param name="history"></param>
        private static void SaveWebhookHistory(WebhookHistoryModel history)
        {

            string storageAccountConnectionString = ConfigurationManager.AppSettings["AppStorage"];
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(storageAccountConnectionString);

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
            CloudTable webhookTable = tableClient.GetTableReference("webhookhistory");

            webhookTable.CreateIfNotExists();

            history.PartitionKey = "ListHistory";
            history.RowKey = history.ListId.ToString();

            TableOperation insertOperation = TableOperation.InsertOrReplace(history);

            TableResult retrievedResult = webhookTable.Execute(insertOperation);

        }
    }
}
