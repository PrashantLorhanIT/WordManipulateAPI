using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Web.Http.Results;

using WordManipulateAPI.Models.EPFM;
using Emc.Documentum.FS.DataModel.Core;
using Emc.Documentum.FS.DataModel.Core.Content;
using System.Configuration;
using System.IO;
using System.Security.Claims;
using System.Web.Http.Cors;
using Newtonsoft.Json;


namespace WordManipulateAPI.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*", SupportsCredentials = true)]
    [RoutePrefix("api/EPFM")]
    //[Authorize]
    public class EPFMController : ApiController
    {
        //// GET api/values
        [HttpGet]
        [Route("GetCabinets")]
        public IEnumerable<CabinetModel> GetCabinets()
        {
            string result = ""; // ActionContext.Request.Headers.Authorization.ToString();
            IEnumerable<CabinetModel> cabinets = null;
            try
            {
                String repository = ConfigurationManager.AppSettings["EPFMRepository"];
                String userName = ConfigurationManager.AppSettings["EPFMUsername"];
                String password = ConfigurationManager.AppSettings["EPFMPassword"];
                String address = ConfigurationManager.AppSettings["EPFMAddress"];

                //String repository = "er_dev_epfm_01";
                //String userName = "svc.cmsadmindev";
                //String password = "Fc$YqK96d8%C";
                //String address = "http://er-kdc-ddcpas01:8082/dfs/services";

                result = "Setting Context";
                MyQueryService t = new MyQueryService();
                t.setContext(userName, password, address, repository);
                cabinets = t.callQueryServiceCabinet();
            }
            catch (Exception ex)
            {
                result = result + "Got exception" + ex.StackTrace;
                throw new Exception(result);
            }

            return cabinets;
        }


        [HttpGet]
        [Route("GetDocuments")]
        public IEnumerable<CabinetModel> GetDocuments(string cabinetid)
        {
            string result = "load";
            IEnumerable<CabinetModel> cabinets = null;
            try
            {
                String repository = ConfigurationManager.AppSettings["EPFMRepository"];
                String userName = ConfigurationManager.AppSettings["EPFMUsername"];
                String password = ConfigurationManager.AppSettings["EPFMPassword"];
                String address = ConfigurationManager.AppSettings["EPFMAddress"];

                //String repository = "er_dev_epfm_01";
                //String userName = "svc.cmsadmindev";
                //String password = "Fc$YqK96d8%C";
                //String address = "http://er-kdc-ddcpas01:8082/dfs/services";

                result = "Setting Context";
                MyQueryService t = new MyQueryService();
                t.setContext(userName, password, address, repository);
                cabinets = t.callQueryServiceDocument(cabinetid);
            }
            catch (Exception ex)
            {
                result = result + "Got exception" + ex.StackTrace;
                throw new Exception(result);
            }

            return cabinets;
        }

        [HttpPost]
        [Route("SearchDocuments")]
        public IEnumerable<DocumentModel> SearchDocuments(SearchModel searchModel)
        {
            string result = "load";
            IEnumerable<DocumentModel> documents = null;
            try
            {
                //string username, password;
                //(username, password) =  SecurityHelper.GetCredentials(User.Identity as ClaimsIdentity);

                String repository = ConfigurationManager.AppSettings["EPFMRepository"];
                String username = ConfigurationManager.AppSettings["EPFMUsername"];
                String password = ConfigurationManager.AppSettings["EPFMPassword"];
                String address = ConfigurationManager.AppSettings["EPFMAddress"];




                Logger.WriteLog("SearchDocuments Log details " + username + Environment.NewLine + password + Environment.NewLine + address);

                SearchServiceDemo searchServiceDemo = new SearchServiceDemo(repository, null, username, password, address);
                Logger.WriteLog("SearchDocuments searchServiceDemo " + searchServiceDemo.ToString());
                documents = searchServiceDemo.SimplePassthroughQueryDocumentWithPath_New(searchModel);
            }
            catch (Exception ex)
            {
                result = result + "Got exception" + ex.StackTrace;
                Logger.WriteLog("SearchDocuments searchServiceDemo " + result);
                throw new Exception(result);
            }

            return documents;
        }

        [HttpGet]
        [Route("DownloadDocument")]
        public IHttpActionResult DownloadDocument(string object_id, int RidType, string category, string SaveFilename, int attachsequence, string attachmentguid)
        {
            string result = "";
            try
            {
                string username, password;
                (username, password) = SecurityHelper.GetCredentials(User.Identity as ClaimsIdentity);
                String repository = ConfigurationManager.AppSettings["EPFMRepository"];
                //String username = ConfigurationManager.AppSettings["EPFMUsername"];
                //String password = ConfigurationManager.AppSettings["EPFMPassword"];
                String address = ConfigurationManager.AppSettings["EPFMAddress"];

             


                ObjectServiceDemo objectService = new ObjectServiceDemo(repository, null, username, password, address);
                //string object_id = "090181cd80054c37";
                ObjectIdentity objIdentity =
                               new ObjectIdentity(new Qualification("dm_document where r_object_id = '" + object_id + "'"),
                                                  repository);

                FileInfo fileInfo = null;
                try
                {
                     fileInfo = objectService.GetWithContent(objIdentity, "Pleasanton", Emc.Documentum.FS.DataModel.Core.Content.ContentTransferMode.MTOM);
                }
                catch(Exception ex)
                { 
                }
                if (fileInfo == null)
                {
                    string TemplateFilePath = ConfigurationManager.AppSettings["TemplateFilePath"];
                    string Fullpath = Path.Combine(TemplateFilePath, SaveFilename);
                    StreamWriter sw = new StreamWriter(Fullpath, false);
                    sw.WriteLine("This is a sample file created at " + DateTime.Now);
                    sw.Close();
                    fileInfo = new FileInfo(Fullpath);
                }
                MoveEPFMResultWrapper resultObj = null;
                //
                using (var client = new HttpClient())
                {
                    try
                    {
                        string uri = ConfigurationManager.AppSettings["CoreMoveEPFMUri"] + "?RidType=" + RidType + "&category=" + 
                                    category + "&Tmpfilename=" + fileInfo.FullName + "&SaveFileName=" + SaveFilename + "&attachsequence=" + attachsequence + "&objectid=" + object_id + "&attachmentguid=" + attachmentguid;
                     
                        client.DefaultRequestHeaders.Authorization = ActionContext.Request.Headers.Authorization;
                        var responseTask = client.GetAsync(uri);
                        responseTask.Wait();

                        var httpresult = responseTask.Result;
                        if (httpresult.IsSuccessStatusCode)
                        {
                            //var readTask = result.Content.ReadAsAsync<IList<StudentViewModel>>();
                            var readTask = httpresult.Content.ReadAsAsync<MoveEPFMResultWrapper>();
                            readTask.Wait();

                            resultObj = readTask.Result;
                            Logger.WriteLog(JsonConvert.SerializeObject(resultObj));

                        }
                        else 
                        {
                            new Exception("ERCMS Core API failed" + httpresult.ReasonPhrase);
                            
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteLog("ERCMS Core API failed " + ex.Message + Environment.NewLine + ex.StackTrace);
                        throw ex;
                    }

                }
                return Content(HttpStatusCode.OK, resultObj);
            }
            catch (Exception ex)
            {
                result = result + "Got exception" + ex.StackTrace;
                return Ok(ex.Message);
            }


        }

        [HttpGet]
        [Route("GetSearchKeywords")]
        public IEnumerable<KeywordModel> GetSearchKeywords()
        {
            string result = "load";
            IEnumerable<KeywordModel> keywords = null;
            try
            {
                string username, password;
                (username, password) = SecurityHelper.GetCredentials(User.Identity as ClaimsIdentity);

                String repository = ConfigurationManager.AppSettings["EPFMRepository"];
                //String username = ConfigurationManager.AppSettings["EPFMUsername"];
                //String password = ConfigurationManager.AppSettings["EPFMPassword"];
                String address = ConfigurationManager.AppSettings["EPFMAddress"];



                //SearchServiceDemo searchServiceDemo = new SearchServiceDemo(repository, null, userName, password, address);
                //keywords = searchServiceDemo.SimplePassthroughQueryForKeywords();
                //result = "Setting Context";

                QueryServiceDemo t = new QueryServiceDemo(repository, null , username, password, address);
                keywords = t.callQueryServiceKeyword();

            }
            catch (Exception ex)
            {
                result = result + "Got exception" + ex.StackTrace;
                throw new Exception(result);
            }

            return keywords;
        }


        // For Archive methods, please refer ObjectServiceDemo.cs under Models folder and look for region named 'CMS Methods'
        //These methods were provided by Shashi, we need to call them logically whenever any folder needs to be created in EPFM under any cabinet OR upload 
        //any file inside any folder.
    }
}
