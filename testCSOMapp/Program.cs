
using OfficeDevPnP.Core.Entities;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Expressions;
using Microsoft.SharePoint.Client.Utilities;
using MTNGW.Core;
using System.Xml.Serialization;
using System.IO;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;

namespace minttulip.spo.remoteprovisioning
{
    class Program
    {        
        static void Main(string[] args)
        {
            var configSites = (NameValueCollection)ConfigurationManager.GetSection("Sites");
            var configLists = (NameValueCollection)ConfigurationManager.GetSection("Lists");            
            Uri siteUri = new Uri(configSites.GetValues("siterequests")[0]);
            string siteRequestListName = configLists.GetValues("siterequestlistname")[0];
            string siteRequestListView = configLists.GetValues("siterequestlistnewitemsview")[0];
            string erroredRequestsListView = configLists.GetValues("siterequestlisterroreditemsview")[0];
            Console.Out.WriteLine("Running as version 1.1");            
            //new requests
            //processNewSiteRequests(siteUri, siteRequestListName, siteRequestListView);            
            //deal with errored requests
            processNewSiteRequests(siteUri, siteRequestListName, erroredRequestsListView);                                    
            Console.Out.WriteLine("done");            
        }

        private static void GetSiteProvisioningTemplate(string siteUrl)
        {
            //Get the realm for the URL
            Uri siteUri = new Uri(siteUrl);
            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);

            //Get the access token for the URL.  
            //   Requires this app to be registered with the tenant
            string accessToken = TokenHelper.GetAppOnlyAccessToken(
                TokenHelper.SharePointPrincipal,
                siteUri.Authority, realm).AccessToken;

            //Get client context with access token
            using (var clientContext =
                TokenHelper.GetClientContextWithAccessToken(
                    siteUri.ToString(), accessToken))
            {
                ProvisioningTemplate newTemplate = clientContext.Web.GetProvisioningTemplate();
                string templateName = clientContext.Web.ServerRelativeUrl.ToLower().Replace("/sites/", "") + ".xml";
                XMLFileSystemTemplateProvider xmlFileProvider = new XMLFileSystemTemplateProvider(@"C:\Users\robmo_000\Documents\xmltemplates", "");
                xmlFileProvider.SaveAs(newTemplate, templateName);
            }
        }

        private static void AddWebPartToPage(Uri siteUri, string pageServerRelUrl, string webPartXml)
        {
            //Get the realm for the URL
            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);

            //Get the access token for the URL.  
            //   Requires this app to be registered with the tenant
            string accessToken = TokenHelper.GetAppOnlyAccessToken(
                TokenHelper.SharePointPrincipal,
                siteUri.Authority, realm).AccessToken;

            //Get client context with access token
            using (var clientContext =
                TokenHelper.GetClientContextWithAccessToken(
                    siteUri.ToString(), accessToken))
            {
                Microsoft.SharePoint.Client.File page = clientContext.Web.GetFileByServerRelativeUrl(pageServerRelUrl);
                page.CheckOut();
                LimitedWebPartManager wpmgr = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
                WebPartDefinition wpd = wpmgr.ImportWebPart(webPartXml);
                wpmgr.AddWebPart(wpd.WebPart, "Left", 0);
                page.CheckIn(String.Empty, CheckinType.MajorCheckIn);                
                clientContext.ExecuteQuery();
            }
        }

        private static void GetWebpartsOnPage(Uri siteUri)
        {
            //Get the realm for the URL
            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);

            //Get the access token for the URL.  
            //   Requires this app to be registered with the tenant
            string accessToken = TokenHelper.GetAppOnlyAccessToken(
                TokenHelper.SharePointPrincipal,
                siteUri.Authority, realm).AccessToken;

            //Get client context with access token
            using (var clientContext =
                TokenHelper.GetClientContextWithAccessToken(
                    siteUri.ToString(), accessToken))
            {


                Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl("/sites/robtestxmlconfig4/default.aspx");
                LimitedWebPartManager wpm = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                clientContext.Load(wpm.WebParts,
                wps => wps.Include(
                wp => wp.WebPart.Title,
                wp => wp.WebPart.TitleUrl,
                wp => wp.WebPart.ZoneIndex));
                clientContext.ExecuteQuery();
                foreach (WebPartDefinition wpd in wpm.WebParts)
                {
                    Microsoft.SharePoint.Client.WebParts.WebPart wp = wpd.WebPart;
                    Console.WriteLine("{0} {1} {2}", wp.Title, wp.TitleUrl, wp.ZoneIndex);                    
                } 
            }
        }        

        private static void processNewSiteRequests(Uri siteUri, string siteRequestListName, string siteRequestListView)
        {
            //Get the realm for the URL
            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);

            //Get the access token for the URL.  
            //   Requires this app to be registered with the tenant
            string accessToken = TokenHelper.GetAppOnlyAccessToken(
                TokenHelper.SharePointPrincipal,
                siteUri.Authority, realm).AccessToken;

            //Get client context with access token
            using (var clientContext =
                TokenHelper.GetClientContextWithAccessToken(
                    siteUri.ToString(), accessToken))
            {
                //do something here                    
                Console.Out.WriteLine("Connected to: {0}, {1}", clientContext.Url, clientContext.ApplicationName);
                var newSiteRequests = GetSiteRequests(clientContext, clientContext.Url, siteRequestListName, siteRequestListView);
                if (newSiteRequests.Count > 0)
                {
                    foreach (SiteProvisioningSiteRequest sr in newSiteRequests)
                    {
                        Console.Out.WriteLine("Creating new site... {0}, {1}, {2}, {3}", sr.Title, sr.Description, sr.SiteTemplate, sr.SiteAdmins[0].LookupValue);
                        //first set site creation as "in progress"
                        UpdateListItem(clientContext, sr.SpItemId, siteRequestListName, "Site_x0020_status", "Being created");
                        //sr.Url = CreateSiteCollection(clientContext, clientContext.Url, sr.Url, sr.SiteTemplate, sr.Title, sr.Description, sr.SiteAdmins, sr.SiteOwners, sr.SiteMembers, sr.SiteVisitors);
                        sr.SiteQuota = 5000;
                        sr.Url = "https://capita.sharepoint.com/sites/" + sr.Url;
                        if (sr.Url != "###ERROR###")
                        {
                            Console.Out.WriteLine("New site created at: " + sr.Url);
                            //now add users to the right groups
                            AddAdditionalSiteAdmins(sr);
                            AddUsersToSiteGroups(sr);
                            UpdateListItem(clientContext, sr.SpItemId, siteRequestListName, "Site_x0020_status", "Available");                            
                            //send email notifying Owners of site readiness
                            SendReadyEmail(clientContext, sr);
                        }
                    }
                }
                else
                {
                    Console.Out.WriteLine("no site requests pending. Exiting...");
                }
            }
        }

        private static void AddAdditionalSiteAdmins(SiteProvisioningSiteRequest sr)
        {
            //Get all groups in site
            Uri siteUri = new Uri(sr.Url);
            //Get the realm for the URL
            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            //Get the access token for the URL.              
            string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;
            //Get client context with access token
            using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken))
            {
                Microsoft.SharePoint.Client.User spUser = null;
                if (sr.SiteAdmins != null)
                {
                    List<UserEntity> adminList = new List<UserEntity>();
                    foreach (FieldUserValue fuv in sr.SiteAdmins)
                    {
                        spUser = clientContext.Web.EnsureUser(fuv.LookupValue);
                        clientContext.Load(spUser);
                        clientContext.ExecuteQuery();
                        adminList.Add(new UserEntity { Email = spUser.Email, LoginName = spUser.LoginName, Title = spUser.Title });
                    }
                    clientContext.Web.AddAdministrators(adminList);
                }
            }
        }

        private static void AddUsersToSiteGroups(SiteProvisioningSiteRequest sr)
        {
            if (sr.SiteOwners != null)
            {
                foreach (FieldUserValue fuv in sr.SiteOwners)
                {
                    AddUserToGroup(fuv.LookupValue, "owners", sr.Url);
                }
            }
            if (sr.SiteMembers != null)
            {
                foreach (FieldUserValue fuv in sr.SiteMembers)
                {
                    AddUserToGroup(fuv.LookupValue, "members", sr.Url);
                }
            }
            if (sr.SiteVisitors != null)
            {
                foreach (FieldUserValue fuv in sr.SiteVisitors)
                {
                    AddUserToGroup(fuv.LookupValue, "visitors", sr.Url);
                }
            }
        }

        private static void SetSitePolicy(string siteUrl, string sitePolicyName)
        {
            //Get all groups in site
            Uri siteUri = new Uri(siteUrl);
            //Get the realm for the URL
            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            //Get the access token for the URL.              
            string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;
            //Get client context with access token
            using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken))
            {
                var policyApplied = false;
                var siteHasPolicy = clientContext.Web.HasSitePolicyApplied();                
                if (!siteHasPolicy)
                {
                    List<SitePolicyEntity> sitePolicies = clientContext.Web.GetSitePolicies();
                    foreach(SitePolicyEntity s in sitePolicies)
                    {                        
                        if (s.Name == sitePolicyName)
                        {
                            clientContext.Web.ApplySitePolicy(s.Name);
                            policyApplied = true;
                        }
                    }
                }
                if (!policyApplied && !siteHasPolicy) { Console.Out.WriteLine("Error: Couldn't apply policy from template. Check propogation from Content Type Hub and sitetypes list configxml."); }
            }
        }

        private static void SendReadyEmail(ClientContext clientContext, SiteProvisioningSiteRequest sr)
        {
            StringBuilder messageBody = new StringBuilder();
            messageBody.Append("<p style=\"font-family: Segoe UI;\">Dear Site Owner,<p>");
            messageBody.Append("<p style=\"font-family: Segoe UI;\">Thank you for using the Capita self-service site creation facility, your site is now ready for use and can be accessed through the following URL:<br/><br/>");
            messageBody.AppendFormat("{0}</p>", sr.Url);
            messageBody.AppendFormat("<p style=\"font-family: Segoe UI;\">Your site is configured in the following way:</p><ul style=\"font-family: Segoe UI;\"><li>Site title: {0}</li><li>Site template: {1}</li><li>Site quota: {2}</li></ul>", sr.Title, sr.SiteTemplate, sr.SiteQuota / 1000 + "GB");
            messageBody.Append("<p style=\"font-family: Segoe UI;\">If you require additional information about this service and what you can expect from your new site, please visit the <a href=\"https://capita.sharepoint.com/sites/help\">Digital Village</a>.<br/><br/>");
            messageBody.Append("*Please do not reply to this email, it is sent from an unmonitored account*</p>");

            EmailProperties props = new EmailProperties();
            List<string> owners = new List<string>();            
            foreach (FieldUserValue fuv in sr.SiteOwners)
            {
                owners.Add(fuv.LookupValue);
            }
            props.To = owners.ToArray();
            props.Subject = "Your new site '" + sr.Title + "' is now available for use";
            props.Body = messageBody.ToString();            
            Microsoft.SharePoint.Client.Utilities.Utility.SendEmail(clientContext, props);
            clientContext.ExecuteQuery();
        }

        private static void UpdateListItem(ClientContext clientContext, int itemId, string listName, string fieldName, string newFieldValue)
        {
            //get list to update
            List updateList = clientContext.Web.Lists.GetByTitle(listName);
            // Get item to update 
            ListItem listItem = updateList.GetItemById(itemId);
            // Update value in specified field
            listItem[fieldName] = newFieldValue;
            listItem.Update();
            clientContext.ExecuteQuery();  
        }

        private static List<SiteProvisioningSiteRequest> GetSiteRequests(ClientContext clientContext, string hostWebUrl, string listName, string viewName)
        {
            //get list items by view
            List<SiteProvisioningSiteRequest> srs = new List<SiteProvisioningSiteRequest>();
            List siteReqs = clientContext.Web.Lists.GetByTitle(listName);
            //load view query
            Microsoft.SharePoint.Client.View siteReqsView = siteReqs.Views.GetByTitle(viewName);
            clientContext.Load(siteReqsView);
            clientContext.ExecuteQuery();
            CamlQuery qry = new CamlQuery();
            qry.ViewXml = "<View><Query>" + siteReqsView.ViewQuery + "</Query></View>";
            //loads list items in view
            ListItemCollection items = siteReqs.GetItems(qry);
            var fieldNames = new[] { "Id", "Title", "Site_x0020_Description", "Site_x0020_Url", "Site_x0020_Admins", "Site_x0020_Owners", "Site_x0020_Members", "Site_x0020_Visitors", "Site_x0020_Template", "Site_x0020_status" };
            List<Expression<Func<ListItemCollection, object>>> allIncludes = new List<Expression<Func<ListItemCollection, object>>>();
            foreach( var f in fieldNames )
            {
                // Log.LogMessage( "Fetching column {0}", c );
                allIncludes.Add( fielditems => fielditems.Include( item => item[ f ] ) );
            }
            clientContext.Load(items, allIncludes.ToArray());
            clientContext.ExecuteQuery();
            if (items.Count > 0)
            {
                var listEnumerator = items.GetEnumerator();
                while(listEnumerator.MoveNext()) {
                    SiteProvisioningSiteRequest sprs = new SiteProvisioningSiteRequest();
                    sprs.Title = listEnumerator.Current["Title"].ToString();
                    sprs.Url = listEnumerator.Current["Site_x0020_Url"].ToString();
                    sprs.Description = listEnumerator.Current["Site_x0020_Description"].ToString();
                    sprs.SiteTemplate = listEnumerator.Current["Site_x0020_Template"].ToString();
                    sprs.SiteStatus = listEnumerator.Current["Site_x0020_status"].ToString();
                    sprs.SiteAdmins = listEnumerator.Current["Site_x0020_Admins"] != null ? (FieldUserValue[])listEnumerator.Current["Site_x0020_Admins"] : null;
                    sprs.SiteOwners = listEnumerator.Current["Site_x0020_Owners"] != null ? (FieldUserValue[])listEnumerator.Current["Site_x0020_Owners"] : null;
                    sprs.SiteMembers = listEnumerator.Current["Site_x0020_Members"] != null ? (FieldUserValue[])listEnumerator.Current["Site_x0020_Members"] : null;
                    sprs.SiteVisitors = (FieldUserValue[])listEnumerator.Current["Site_x0020_Visitors"];
                    sprs.SpItemId = listEnumerator.Current.Id;
                    srs.Add(sprs);                    
                }                
            }
            return srs;
        }

        /// <summary>
        /// Creates a new site collection using the parameters specified
        /// </summary>
        /// <param name="hostWebUrl"></param>
        /// <param name="url"></param>
        /// <param name="template"></param>
        /// <param name="title"></param>
        /// <param name="adminAccount"></param>
        /// <returns></returns>
        private static string CreateSiteCollection(ClientContext ctx, string hostWebUrl, SiteProvisioningSiteRequest sr)
        {
            //get the base tenant admin urls
            var tenantStr = hostWebUrl.ToLower().Replace("-my", "").Substring(8);
            tenantStr = tenantStr.Substring(0, tenantStr.IndexOf("."));

            //get the current user to set as owner
            Microsoft.SharePoint.Client.User ownerUser = null;
            if (sr.SiteAdmins != null && sr.SiteAdmins[0] != null)
            {
                string emailtouse = sr.SiteAdmins[0].LookupValue;// "i:0#.f|membership|p10260756@capita.co.uk";
                ownerUser = ctx.Web.EnsureUser(emailtouse);
            }
            if (ownerUser != null)
            {
                ctx.Load(ownerUser);
                ctx.ExecuteQuery();

                //create site collection using the Tenant object
                var webUrl = String.Format("https://{0}.sharepoint.com/{1}/{2}", tenantStr, "sites", sr.Url);
                var tenantAdminUri = new Uri(String.Format("https://{0}-admin.sharepoint.com", tenantStr));
                string realm = TokenHelper.GetRealmFromTargetUrl(tenantAdminUri);
                string token = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, tenantAdminUri.Authority, realm).AccessToken;
                SiteProvisioningSiteRequestType srt = new SiteProvisioningSiteRequestType();
                //Get client context with access token
                using (var adminContext = TokenHelper.GetClientContextWithAccessToken(tenantAdminUri.ToString(), token))
                {
                    srt = GetSiteTypeInfo(sr.SiteTemplate, ctx);
                    var tenant = new Tenant(adminContext);
                    var primaryAdminEmail = ownerUser.Email;
                    if (!ownerUser.LoginName.Contains(primaryAdminEmail))
                    {
                        //set owner email to proper tenant value, rather than masked, email so it gets recognised
                        try
                        {
                            primaryAdminEmail = ownerUser.LoginName.Split('|')[2];
                        }
                        catch { primaryAdminEmail = ownerUser.Email; }
                    }
                    sr.SiteQuota = srt.SiteStorageQuota;
                    var properties = new SiteCreationProperties()
                    {
                        Url = webUrl,
                        Owner = primaryAdminEmail,
                        Title = sr.Title,
                        Template = srt.TypeSharePointSiteId,
                        StorageMaximumLevel = srt.SiteStorageQuota, //MB storage available
                        UserCodeMaximumLevel = srt.SiteResourceLevel,     
                        TimeZoneId = 2
                    };

                    //start the SPO operation to create the site
                    SpoOperation op = tenant.CreateSite(properties);
                    adminContext.Load(tenant);
                    adminContext.Load(op, i => i.IsComplete);
                    adminContext.ExecuteQuery();

                    //check if site creation operation is complete
                    while (!op.IsComplete)
                    {
                        //wait 30seconds and try again
                        System.Threading.Thread.Sleep(30000);
                        op.RefreshLoad();
                        adminContext.ExecuteQuery();
                    }
                }

                //setup the new site collection
                SetSiteTheme(webUrl, realm, ref token, srt);
                CreateSiteLibraries(webUrl, realm, ref token, srt);
                SetSiteRegionAndLocale(webUrl, realm, ref token, srt);
                SetSitePolicy(webUrl, srt.DefaultSitePolicyName);
                //return site url
                return webUrl;
            }
            else
            {
                return "###ERROR###";
            }
        }        

        private static void SetSiteRegionAndLocale(string webUrl, string realm, ref string token, SiteProvisioningSiteRequestType srt)
        {
            var siteUri = new Uri(webUrl);
            token = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;
            using (var newWebContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), token))
            {
                Web web = newWebContext.Web;
                RegionalSettings regSettings = web.RegionalSettings;                
                newWebContext.Load(web);
                newWebContext.Load(regSettings);
                newWebContext.ExecuteQuery();                
                newWebContext.Web.RegionalSettings.LocaleId = uint.Parse("2057");                
                newWebContext.Web.RegionalSettings.Update();
                newWebContext.ExecuteQuery();
            }
        }       

        private static void CreateSiteLibraries(string webUrl, string realm, ref string token, SiteProvisioningSiteRequestType srt)
        {
            foreach (ProvisionedLibrary pl in srt.Libraries)
            {
                var siteUri = new Uri(webUrl);
                token = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;
                using (var newWebContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), token))
                {
                    var listCreated = false;
                    newWebContext.Load(newWebContext.Web.Lists, lists => 
                        lists.Include( 
                            list => list.Title,
                            list => list.Id
                        )
                    );
                    // Execute query. 
                    newWebContext.ExecuteQuery();
                    foreach (List spl in newWebContext.Web.Lists)
                    {
                        if (spl.Title == pl.Title)
                        {
                            listCreated = true;
                        }
                    }
                    if (!listCreated) {
                        AddNewListToSite(pl, newWebContext);
                    }
                    else {
                        AddContentTypesToList(pl, newWebContext);
                    }                 
                }
            }            
        }

        private static void AddContentTypesToList(ProvisionedLibrary pl, ClientContext newWebContext)
        {
            //list already exists just do Content type check
            List list = newWebContext.Web.Lists.GetByTitle(pl.Title);
            ContentTypeCollection ctColl = list.ContentTypes;
            newWebContext.Load(list);
            newWebContext.Load(ctColl);
            newWebContext.ExecuteQueryRetry();
            List<string> contentTypesInList = new List<string>();
            foreach (Microsoft.SharePoint.Client.ContentType ct in ctColl)
            {
                contentTypesInList.Add(ct.Id.StringValue);
                contentTypesInList.Add(ct.Name);
            }
            foreach (ProvisionedContentType pct in pl.ContentTypes)
            {
                try
                {
                    if (!contentTypesInList.Contains(pct.ContentTypeId) && !contentTypesInList.Contains(pct.Name))
                    {
                        //add it
                        Microsoft.SharePoint.Client.ContentType ct = newWebContext.Web.ContentTypes.GetById(pct.ContentTypeId);
                        newWebContext.Load(ct);
                        newWebContext.ExecuteQuery();
                        list.ContentTypes.AddExistingContentType(ct);
                        list.Update();
                        newWebContext.ExecuteQuery();
                    }
                }
                catch (Exception ex) { Console.Out.WriteLine("Couldn't add content type to library because {0}", ex.Message); }
            }
        }

        private static void AddNewListToSite(ProvisionedLibrary pl, ClientContext newWebContext)
        {
            //create the list
            ListCreationInformation newListInfo = new ListCreationInformation()
            {
                Title = pl.Title,
                TemplateType = pl.TemplateType                
            };
            List newList = newWebContext.Web.Lists.Add(newListInfo);
            newList.ContentTypesEnabled = true;
            newList.Update();
            newWebContext.ExecuteQuery();
            //now check content types
            List<string> contentTypesInList = new List<string>();
            ContentTypeCollection ctColl = newList.ContentTypes;
            newWebContext.Load(newList);
            newWebContext.Load(ctColl);
            newWebContext.ExecuteQuery();
            foreach (Microsoft.SharePoint.Client.ContentType ct in ctColl)
            {
                contentTypesInList.Add(ct.Id.StringValue);
                contentTypesInList.Add(ct.Name);
            }
            foreach (ProvisionedContentType pct in pl.ContentTypes)
            {
                try
                {
                    if (!contentTypesInList.Contains(pct.ContentTypeId) && !contentTypesInList.Contains(pct.Name))
                    {
                        //add it
                        Microsoft.SharePoint.Client.ContentType ct = newWebContext.Web.ContentTypes.GetById(pct.ContentTypeId);
                        newWebContext.Load(ct);
                        newWebContext.ExecuteQuery();
                        newList.ContentTypes.AddExistingContentType(ct);
                        newList.Update();
                        newWebContext.ExecuteQuery();
                    }
                }
                catch (Exception ex) { Console.Out.WriteLine("Couldn't add content type to library because {0}", ex.Message); }
            }
        } 

        private static void SetSiteTheme(string webUrl, string realm, ref string token, SiteProvisioningSiteRequestType srt)
        {
            var siteUri = new Uri(webUrl);
            token = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;
            using (var newWebContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), token))
            {
                var newWeb = newWebContext.Web;
                newWebContext.Load(newWeb);
                newWebContext.ExecuteQuery();
                new LabHelper().SetThemeBasedOnName(newWebContext, newWeb, newWeb, srt.ThemeColourName);
                newWeb.SiteLogoUrl = srt.CustomLogoUrl;
                newWeb.AlternateCssUrl = srt.CustomCssUrl;
                newWeb.Update();
                newWebContext.ExecuteQuery();
            }
        }

        private static SiteProvisioningSiteRequestType GetSiteTypeInfo(string siteTypeName, ClientContext clientContext)
        {
            var configLists = (NameValueCollection)ConfigurationManager.GetSection("Lists");
            var listName = configLists.GetValues("siteconfiglistname")[0];
            //get list items by view
            SiteProvisioningSiteRequestType srt = new SiteProvisioningSiteRequestType();
            List siteReqs = clientContext.Web.Lists.GetByTitle(listName);                       
            CamlQuery qry = new CamlQuery();
            qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + siteTypeName + "</Value></Eq></Where></Query></View>";
            ListItemCollection items = siteReqs.GetItems(qry);
            var fieldNames = new[] { "Id", "Title", "versionnumber", "configxml", "datelastupdated" };
            List<Expression<Func<ListItemCollection, object>>> allIncludes = new List<Expression<Func<ListItemCollection, object>>>();
            foreach (var f in fieldNames)
            {
                // Log.LogMessage( "Fetching column {0}", c );
                allIncludes.Add(fielditems => fielditems.Include(item => item[f]));
            }
            clientContext.Load(items, allIncludes.ToArray());
            clientContext.ExecuteQuery();
            if (items.Count == 1)
            {
                var listEnumerator = items.GetEnumerator();
                while (listEnumerator.MoveNext())
                {
                    if(listEnumerator.Current["Title"].ToString() == siteTypeName)
                    {
                        //unpack the xml
                        var xmlDef = listEnumerator.Current["configxml"].ToString();
                        srt = GetSrtFromXml(xmlDef);
                        break;
                    }                    
                }
            }
            return srt;
        }

        private static SiteProvisioningSiteRequestType GetSrtFromXml(string srcXml)
        {
            SiteProvisioningSiteRequestType srtRet;
            var serializer = new XmlSerializer(typeof(SiteProvisioningSiteRequestType));
            using (TextReader tr = new StringReader(srcXml))
            {
                srtRet = (SiteProvisioningSiteRequestType)serializer.Deserialize(tr);
            }
            return srtRet;
        }

        private static bool AddUserToGroup(string userEmail, string groupType, string siteUrl)
        {
            //Get all groups in site
            Uri siteUri = new Uri(siteUrl);
            //Get the realm for the URL
            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            //Get the access token for the URL.              
            string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;

            //Get client context with access token
            using (var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken))
            {
                var groupCollection = clientContext.Web.SiteGroups;
                clientContext.Load(groupCollection);
                clientContext.ExecuteQuery();
                foreach (Group g in groupCollection)
                {
                    if (g.Title.ToLower().Contains(groupType))
                    {
                        //get users, add current to group
                        Microsoft.SharePoint.Client.User userToAdd = clientContext.Web.EnsureUser(userEmail);
                        var usersInGroup = g.Users;
                        usersInGroup.AddUser(userToAdd);
                        clientContext.Load(usersInGroup);
                        clientContext.ExecuteQuery();
                        return true;
                    }
                }
            }            
            return false;
        }


        
        private static string URLCombine(string baseUrl, string relativeUrl)
        {
            if (baseUrl.Length == 0)
                return relativeUrl;
            if (relativeUrl.Length == 0)
                return baseUrl;
            return string.Format("{0}/{1}",
                baseUrl.TrimEnd(new char[] { '/', '\\' }),
                relativeUrl.TrimStart(new char[] { '/', '\\' }));
        }

        private static void GenerateDummyXml()
        {
            ProvisionedContentType ct1 = new ProvisionedContentType { ContentTypeId = "0x0101", Name = "Document" };
            ProvisionedContentType ct2 = new ProvisionedContentType { ContentTypeId = "0x0105", Name = "Link" };
            ProvisionedContentType ct3 = new ProvisionedContentType { ContentTypeId = "0x0106", Name = "Contact" };
            List<ProvisionedContentType> cts = new List<ProvisionedContentType> { ct1, ct2, ct3 };
            ProvisionedLibrary lb1 = new ProvisionedLibrary { TemplateType = 101, ContentTypes = cts, Title = "Team Shared Documents" };
            ProvisionedLibrary lb2 = new ProvisionedLibrary { TemplateType = 101, ContentTypes = cts, Title = "Team Private Documents" };
            ProvisionedLibrary lb3 = new ProvisionedLibrary { TemplateType = 101, ContentTypes = cts, Title = "Team External Documents" };
            List<ProvisionedLibrary> lbs = new List<ProvisionedLibrary> { lb1, lb2, lb3 };
            SiteProvisioningSiteRequestType srt1 = new SiteProvisioningSiteRequestType
            {
                TypeName = "Team Site",
                Libraries = lbs,
                ThemeColourName = "Orange",
                SiteStorageQuota = 5000,
                SiteResourceLevel = 100,
                CustomCssUrl = "https://testminttulip.sharepoint.com/sites/standardsource/src/teamsite.css",
                CustomLogoUrl = "https://testminttulip.sharepoint.com/sites/standardsource/src/teamsitelogo.png",
                DefaultSitePolicyName = "Team Site"
            };
            XmlSerializer x = new XmlSerializer(srt1.GetType());
            TextWriter writer = new StreamWriter(@"C:\Users\robmo_000\Documents\sitereq.xml");
            x.Serialize(writer, srt1);
            writer.Close();
        }
    }
}
