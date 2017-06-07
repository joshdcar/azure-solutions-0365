using System;
using System.Linq;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using System.Threading.Tasks;
using System.Configuration;
using Newtonsoft.Json;
using SmartPhotoLibrary.Azure.Functions.Models;
using System.Collections.Generic;
using System.Net.Http;
using System.Web;
using System.Text;
using System.Net.Http.Headers;
using System.IO;

namespace SmartPhotoLibrary.Azure.Functions
{
    public static class PhotoRequests
    {
        private static TraceWriter _log = null;

        [FunctionName("PhotoRequests")]        
        public static async Task Run([QueueTrigger("photorequests", Connection = "AppStorage")]string requestMessage, TraceWriter log)
        {
            try
            {
                _log = log;

                log.Info($"Queue event for photorequests. Message: {requestMessage}");

                var request = JsonConvert.DeserializeObject<PhotoRequestModel>(requestMessage);

                var fileRequest = await GetSharePointPhotoFile(request);

                if (fileRequest != null)
                {
                    var visionResponse = await AnalyzePhoto(fileRequest);

                    UpdateSharePointMetadata(visionResponse, request);
                }
            }
            catch(Exception ex)
            {
                log.Error("An exception occured while processing the photo request. Exception: " + ex.ToString());
            }
            

        }

        /// <summary>
        /// Retrieve the file from SharePoint
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        private static async Task<byte[]> GetSharePointPhotoFile(PhotoRequestModel request)
        {
            _log.Info($"Retrieve SharePoint File");

            byte[] file = null;

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

                var listID = request.ListId;

                IEnumerable<List> listResults = clientContext.LoadQuery<List>(lists.Where(lst => lst.Id == listID));
                clientContext.ExecuteQueryRetry();

                var photoLibrary = listResults.FirstOrDefault();

                if (photoLibrary != null)
                {
                    var query = new CamlQuery();
                    query.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>{request.Id}</Value></Eq></Where></Query><ViewFields><FieldRef Name='FileRef' /><FieldRef Name='FileLeafRef' /></ViewFields></View>";

                    var items = photoLibrary.GetItems(query);

                    clientContext.Load(items, includes => includes.Include(i => i.File,i => i.File.ServerRelativeUrl));
                    clientContext.ExecuteQuery();

                    var fileItem = items.FirstOrDefault();

                    if (fileItem != null)
                    {
                   
                        var spFile = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileItem.File.ServerRelativeUrl);

                        _log.Info($"Retrieve file {fileItem.File.ServerRelativeUrl}");

                        using (var stream = spFile.Stream)
                        {
                            using (var memStream = new MemoryStream())
                            {
                                await stream.CopyToAsync(memStream);
                                return memStream.ToArray();
                            }

                        }
                    }

                }
            }

            return file;
        }

        /// <summary>
        /// Call Microsoft Cognitive Services Vision API to Analyze our Photo
        /// </summary>
        /// <param name="fileBytes"></param>
        /// <returns></returns>
        private static async Task<VisionResponseModel> AnalyzePhoto(byte[] fileBytes)
        {
            _log.Info($"Analyzing Photo");

            var subscriptionKey = ConfigurationManager.AppSettings["computervisionkey"];
            var visionApiUrl = ConfigurationManager.AppSettings["computervisionApiUrl"];

            var client = new HttpClient();
            var queryString = HttpUtility.ParseQueryString(string.Empty);

            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", subscriptionKey);

            var uri = $"{visionApiUrl}/analyze?visualFeatures=Categories,Tags,Adult,Color,Description";

            using (var content = new ByteArrayContent(fileBytes))
            {
                content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

                var response = await client.PostAsync(uri, content);

                string responseBody = await response.Content.ReadAsStringAsync();

                _log.Info($"Analysis Results {responseBody}");

                var visionresult = JsonConvert.DeserializeObject<VisionResponseModel>(responseBody);

                return visionresult;

            }

            
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name=""></param>
        private static void UpdateSharePointMetadata(VisionResponseModel visionModel, PhotoRequestModel request)
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

                var listID = request.ListId;

                IEnumerable<List> listResults = clientContext.LoadQuery<List>(lists.Where(lst => lst.Id == listID));
                clientContext.ExecuteQueryRetry();

                var photoLibrary = listResults.FirstOrDefault();

                if (photoLibrary != null)
                {
                    var query = new CamlQuery();
                    query.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>{request.Id}</Value></Eq></Where></Query><ViewFields><FieldRef Name='FileRef' /><FieldRef Name='FileLeafRef' /></ViewFields></View>";

                    var items = photoLibrary.GetItems(query);

                    clientContext.Load(items);
                    clientContext.ExecuteQuery();

                    var fileItem = items.FirstOrDefault();

                    if (fileItem != null)
                    {
                        //We want to be a little selective with our tags so we're only going to
                        //add tags with a confidence level of greater then 75% (.75)
                        var tags = visionModel.tags.Where(t => t.confidence > .75M).Select(t => t.name).ToArray();
                        var colors = new String[] { visionModel.color.dominantColorBackground, visionModel.color.dominantColorForeground };


                        fileItem["Tags"] = string.Join(",", tags);
                        fileItem["LastAnalyzed"] = DateTime.Now;
                        fileItem["Analyzed"] = true;
                        fileItem["Inappropriate"] = visionModel.adult.isAdultContent || visionModel.adult.isRacyContent;
                        fileItem["Colors"] = colors;

                        fileItem.Update();

                        clientContext.ExecuteQuery();

                        _log.Info($"SharePoint File Updated with tags {string.Join(",", tags)}");
                    }

                }
            }
        }

    }
}
