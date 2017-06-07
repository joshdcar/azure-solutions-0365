using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using SmartPhotoLibrary.Azure.Functions.Models;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using System.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using Microsoft.WindowsAzure.Storage.Table;

namespace SmartPhotoLibrary.Azure.Functions
{
    public static class PhotoLibraryWebhook
    {
        [FunctionName("PhotoLibraryWebhook")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info($"Smart Photo Library Webhook Triggered");

            // Grab the validationToken URL parameter
            string validationToken = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
                .Value;

            // If a validation token is present, we need to respond within 5 seconds by  
            // returning the given validation token. This only happens when a new 
            // web hook is being added
            if (validationToken != null)
            {
                log.Info($"Validation token {validationToken} received");
                var response = req.CreateResponse(HttpStatusCode.OK);
                response.Content = new StringContent(validationToken);
                return response;
            }

            await ProcessWebhookEvent(req, log);

            return new HttpResponseMessage(HttpStatusCode.OK);
        }

        /// <summary>
        /// Process the webhook Notification Event
        /// </summary>
        /// <param name="req"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        private static async Task ProcessWebhookEvent(HttpRequestMessage req, TraceWriter log)
        {
            log.Info($"SharePoint triggered our webhook with a photo change.");
            var content = await req.Content.ReadAsStringAsync();
            log.Info($"Message Content: {content}");

            var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content).Value;
            log.Info($"Found {notifications.Count} notifications");

            if (notifications.Count > 0)
            {
                log.Info($"Processing notifications...");

                foreach (var notification in notifications)
                {
                    await QueueNotificationProcessing(notification);
                }
            }
        }

        /// <summary>
        /// Submit the notification for processing. Webhooks must respond within 5 seconds so we'll queue the notification
        /// instead of processing it in case we take longer then 5 seconds.
        /// </summary>
        /// <param name="changeList"></param>
        /// <param name="change"></param>
        /// <returns></returns>
        private static async Task QueueNotificationProcessing(NotificationModel notification)
        {

            string storageAccountConnectionString = ConfigurationManager.AppSettings["AppStorage"];
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(storageAccountConnectionString);

            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference("photonotifications");

            queue.CreateIfNotExists();

            string message = JsonConvert.SerializeObject(notification);

            await queue.AddMessageAsync(new CloudQueueMessage(message));

        }



    }
}