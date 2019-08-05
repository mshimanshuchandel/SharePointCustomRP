using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SharePointManager.Helpers;
using SharePointManager.Models;

namespace SharePointManager
{
    public static class CustomerSharePointManagement
    {
        static string siteUrl = Environment.GetEnvironmentVariable("spSiteUrl");
        static string webTemplateUrl = Environment.GetEnvironmentVariable("spWebTemplate");
        static int webLanguage = Convert.ToInt32(Environment.GetEnvironmentVariable("spWebLanguage"));

        [FunctionName("sites")]
        public static async Task<HttpResponseMessage> sites([HttpTrigger(AuthorizationLevel.Anonymous, "get", "put", "delete",
            Route = "subscriptions/{subscriptionId}/resourcegroups/{resourceGroupname}/providers/Microsoft.Customproviders/resourceproviders/{providerName}/sites/{url?}")]
            HttpRequestMessage req, string subscriptionId, string resourceGroupName, string providerName, string url, TraceWriter log)
        {
            var realm = TokenHelper.GetRealmFromTargetUrl(new Uri(siteUrl));

            //Get the access token for the URL.  
            var accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, new Uri(siteUrl).Authority, realm).AccessToken;

            //Get method to return a customer site information
            if (req.Method == HttpMethod.Get)
            {
                //All sites
                if (string.IsNullOrEmpty(url))
                {
                    using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUrl, accessToken))
                    {
                        var subWebs = clientContext.Web.Webs;
                        clientContext.Load(subWebs);
                        clientContext.ExecuteQuery();

                        var response = new List<string>();
                        foreach (var subweb in subWebs)
                        {
                            response.Add(subweb.Title);
                        }

                        var jsonResponse = JsonConvert.SerializeObject(response);

                        var responseContent = new
                        {
                            sites = response
                        };
                        return GetPropertyResult(req, HttpStatusCode.OK, responseContent, "sites");
                    }
                }
                else
                {
                    //returns a customer's site information
                    using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUrl, accessToken))
                    {

                        var web = clientContext.Site.OpenWeb(url);

                        clientContext.Load(web);
                        clientContext.ExecuteQuery();

                        var responseContent = new
                        {
                            title = web.Title,
                            description = web.Description,
                            url = web.Url
                        };

                        return GetResourceResult(req, responseContent, url, "sites", subscriptionId, resourceGroupName, providerName, "sites");
                    }
                }
            }
            //Put method to create a new customer site
            else if (req.Method == HttpMethod.Put)
            {
                var content = req.Content;
                var responseString = content.ReadAsStringAsync().Result;
                var jsonContent = JToken.Parse(responseString).SelectToken("properties").ToString();

                var webInfo = JsonConvert.DeserializeObject<WebModel>(jsonContent);

                var properRequest = false;
                var urlInfo = string.Empty;

                if (!string.IsNullOrEmpty(url) || !string.IsNullOrEmpty(webInfo.Url))
                {
                    properRequest = true;
                    urlInfo = !string.IsNullOrEmpty(url) ? url : webInfo.Url;
                }
                if (properRequest)
                {
                    using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUrl, accessToken))
                    {
                        var wci = new WebCreationInformation();

                        wci.Url = urlInfo;
                        wci.Title = webInfo.Title;
                        wci.Description = webInfo.Description;
                        wci.UseSamePermissionsAsParentSite = true;
                        wci.WebTemplate = webTemplateUrl;
                        wci.Language = 1033;

                        Web w = clientContext.Site.RootWeb.Webs.Add(wci);

                        clientContext.ExecuteQuery();

                        var response = "web created";

                        var responseContent = new
                        {
                            message = response,
                            title = webInfo.Title,
                            description = webInfo.Description,
                            url = siteUrl + "/" + urlInfo
                        };

                        return GetResourceResult(req, responseContent, url, "sites", subscriptionId, resourceGroupName, providerName, "sites");
                    }
                }
                else
                {
                    return req.CreateResponse(HttpStatusCode.BadRequest, "Bad request. Need url information.");
                }
            }
            //Delete method to delete a site
            else if (req.Method == HttpMethod.Delete)
            {
                var content = req.Content;
                string JsonContent = content.ReadAsStringAsync().Result;

                using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUrl, accessToken))
                {
                    var web = clientContext.Site.OpenWeb(url);

                    clientContext.Load(web);
                    try
                    {
                        clientContext.ExecuteQuery();
                    }
                    catch(Exception ex)
                    {
                        var badRequest = new
                        {
                            message = "site does not exist"
                        };
                        return GetPropertyResult(req, HttpStatusCode.NotFound, badRequest, "sites");
                    }

                    web.DeleteObject();
                    clientContext.ExecuteQuery();

                    var responseContent = new
                    {
                        message = "site deleted"
                    };
                    return GetPropertyResult(req, HttpStatusCode.OK, responseContent, "sites");
                }

            }
            else
            {
                return req.CreateResponse(HttpStatusCode.BadRequest, "Bad request. Customer site supports only get, put and delete verbs.");
            }

        }

        [FunctionName("events")]
        public static async Task<HttpResponseMessage> events([HttpTrigger(AuthorizationLevel.Anonymous, "get", "put", "delete",
            Route = "subscriptions/{subscriptionId}/resourcegroups/{resourceGroupname}/providers/Microsoft.Customproviders/resourceproviders/{providerName}/sites/{url}/events/{eventName?}")]
            HttpRequestMessage req, string subscriptionId, string resourceGroupName, string providerName, string url, string eventName, TraceWriter log)
        {
            var realm = TokenHelper.GetRealmFromTargetUrl(new Uri(siteUrl));

            //Get the access token for the URL.  
            var accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, new Uri(siteUrl).Authority, realm).AccessToken;

            //Get methods to get all Events created in all webs
            if (req.Method == HttpMethod.Get)
            {
                //All events created in a customer site
                if (string.IsNullOrEmpty(eventName))
                {
                    using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUrl, accessToken))
                    {

                        var subWeb = clientContext.Site.OpenWeb(url);

                        clientContext.Load(subWeb);
                        clientContext.ExecuteQuery();

                        var response = new List<string>();

                        var eventList = subWeb.GetListByTitle("Events");
                        clientContext.Load(eventList);
                        clientContext.ExecuteQuery();

                        CamlQuery camlQuery = new CamlQuery();
                        camlQuery.ViewXml = "<View><RowLimit>100</RowLimit></View>";

                        var itemCollection = eventList.GetItems(camlQuery);

                        clientContext.Load(itemCollection, items => items.Include(
                            item => item.Id,
                            item => item.DisplayName));

                        clientContext.ExecuteQuery();

                        foreach (var item in itemCollection)
                        {
                            response.Add(item.DisplayName);
                        }



                        var responseContent = new
                        {
                            site = subWeb.Title,
                            customerEvents = response
                        };

                        return GetResourceResult(req, responseContent, url, "events", subscriptionId, resourceGroupName, providerName, "events");
                    }
                }
                //Fetch an event in a customer site
                else
                {
                    using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUrl, accessToken))
                    {
                        var subWeb = clientContext.Site.OpenWeb(url);

                        clientContext.Load(subWeb);
                        clientContext.ExecuteQuery();

                        var response = new List<string>();

                        var eventList = subWeb.GetListByTitle("Events");
                        clientContext.Load(eventList);
                        clientContext.ExecuteQuery();

                        CamlQuery camlQuery = new CamlQuery();
                        camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + eventName + "</Value></Eq></Where></Query></View>";

                        var itemCollection = eventList.GetItems(camlQuery);

                        clientContext.Load(itemCollection, items => items.Include(
                    item => item.Id,
                    item => item.DisplayName));

                        clientContext.ExecuteQuery();

                        if (itemCollection.Count > 0)
                        {
                            var responseContent = new
                            {
                                site = subWeb.Title,
                                customerEvent = new
                                {
                                    title = itemCollection[0].DisplayName,
                                    id = itemCollection[0].Id,
                                }
                            };

                            return GetResourceResult(req, responseContent, eventName, "events", subscriptionId, resourceGroupName, providerName, "events");
                        }
                        else
                        {
                            var responseContent = new
                            {
                                message = "item does not exist"
                            };

                            return GetResourceResult(req, responseContent, eventName, "events", subscriptionId, resourceGroupName, providerName, "events");
                        }
                    }
                }


            }
            //Put method to broadcast event to all customers
            else if (req.Method == HttpMethod.Put)
            {
                var content = req.Content;
                var responseString = content.ReadAsStringAsync().Result;
                var jsonContent = JToken.Parse(responseString).SelectToken("properties").ToString();

                var eventInfo = JsonConvert.DeserializeObject<CreateEventModel>(jsonContent);

                using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUrl, accessToken))
                {

                    var subWeb = clientContext.Site.OpenWeb(url);

                    clientContext.Load(subWeb);
                    clientContext.ExecuteQuery();

                    var eventList = subWeb.GetListByTitle("Events");
                    clientContext.Load(eventList);
                    clientContext.ExecuteQuery();

                    clientContext.Load(subWeb);
                    clientContext.ExecuteQuery();

                    var eventsList = subWeb.Lists.GetByTitle("Events");

                    var itemCreateInfo = new ListItemCreationInformation();
                    var newItem = eventsList.AddItem(itemCreateInfo);
                    newItem["Title"] = eventName;
                    newItem["EventDate"] = eventInfo.StartTime;
                    newItem["EndDate"] = eventInfo.EndTime;
                    newItem["Description"] = eventInfo.Description;
                    newItem.Update();

                    clientContext.ExecuteQuery();

                    var response = "event posted";

                    var responseContent = new
                    {
                        site = subWeb.Title,
                        customerEvent = new
                        {
                            message = response,
                            title = eventName,
                            description = eventInfo.Description,
                            startTime = eventInfo.StartTime,
                            endTime = eventInfo.EndTime,
                        }
                    };

                    return GetResourceResult(req, responseContent, eventName, "events", subscriptionId, resourceGroupName, providerName, "events");
                }
            }
            else if (req.Method == HttpMethod.Delete)
            {
                using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUrl, accessToken))
                {
                    var subWeb = clientContext.Site.OpenWeb(url);

                    clientContext.Load(subWeb);
                    clientContext.ExecuteQuery();

                    var response = new List<string>();

                    var eventList = subWeb.GetListByTitle("Events");
                    clientContext.Load(eventList);
                    clientContext.ExecuteQuery();

                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + eventName + "</Value></Eq></Where></Query></View>";

                    var itemCollection = eventList.GetItems(camlQuery);

                    clientContext.Load(itemCollection, items => items.Include(
                item => item.Id,
                item => item.DisplayName));

                    clientContext.ExecuteQuery();

                    if (itemCollection.Count > 0)
                    {
                        itemCollection[0].DeleteObject();
                        clientContext.ExecuteQuery();
                    }
                    else
                    {
                        var badRequest = new
                        {
                            message = "item does not exist"
                        };

                        return GetPropertyResult(req, HttpStatusCode.NotFound, badRequest, "events");
                    }
                    var responseContent = new
                    {
                        message = "event deleted"
                    };
                    return GetPropertyResult(req, HttpStatusCode.OK, responseContent, "events");
                }
            }
            else
            {
                return req.CreateResponse(HttpStatusCode.BadRequest, "Bad request");
            }
        }
        private static HttpResponseMessage GetResourceResult(HttpRequestMessage req, object responseContent, string resourceName, string resourceType, string subscriptionId, string resourceGroupName, string providerName, string resourceTypeName)
        {
            return req.CreateResponse(HttpStatusCode.OK, (new ArmResource(
                resourceName,
                resourceType,
                subscriptionId,
                resourceGroupName,
                providerName,
                resourceTypeName,
                responseContent
            )));
        }

        private static HttpResponseMessage GetPropertyResult(HttpRequestMessage req, HttpStatusCode statusCode, object responseContent, string resourceTypeName)
        {
            return req.CreateResponse(statusCode, (new ArmProperty(responseContent, resourceTypeName)));
        }
    }
}
