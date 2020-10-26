using System;
using System.Collections.Generic;
using System.Text;
using Emc.Documentum.FS;
using Emc.Documentum.FS.Services.Core;
using Emc.Documentum.FS.DataModel.Core;
using Emc.Documentum.FS.DataModel.Core.Query;
using Emc.Documentum.FS.DataModel.Core.Properties;
using Emc.Documentum.FS.Runtime;
using Emc.Documentum.FS.Runtime.Context;

namespace WordManipulateAPI.Models.EPFM
{
    public class QueryServiceDemo : DemoBase
    {
        private IQueryService queryService;

        public QueryServiceDemo(String defaultRepository, String secondaryRepository, String userName, String password, string domain)
            : base(defaultRepository, secondaryRepository, userName, password)
        {
            ServiceFactory serviceFactory = ServiceFactory.Instance;
            queryService =
                serviceFactory.GetRemoteService<IQueryService>(DemoServiceContext, "core", domain);
        }

        public void BasicPassthroughQuery()
        {
            PassthroughQuery query = new PassthroughQuery();
            query.QueryString = "select r_object_id, "
                                + "object_name from dm_cabinet";
            query.AddRepository(DefaultRepository);
            QueryExecution queryEx = new QueryExecution();
            queryEx.CacheStrategyType = CacheStrategyType.DEFAULT_CACHE_STRATEGY;
            OperationOptions operationOptions = null;
            QueryResult queryResult = queryService.Execute(query, queryEx, operationOptions);
            Console.WriteLine("QueryId == " + query.QueryString);
            Console.WriteLine("CacheStrategyType == " + queryEx.CacheStrategyType);
            DataPackage resultDp = queryResult.DataPackage;
            List<DataObject> dataObjects = resultDp.DataObjects;
            Console.WriteLine("Total objects returned is: " + dataObjects.Count);
            foreach (DataObject dObj in dataObjects)
            {
                PropertySet docProperties = dObj.Properties;
                String objectId = dObj.Identity.GetValueAsString();
                String docName = docProperties.Get("object_name").GetValueAsString();
                Console.WriteLine("Document " + objectId + " name is " + docName);
            }
        }

        /*
        *    Sequentially processes a cached query result
        *    Terminates when end of query result is reached
        */
        //public IEnumerable<DocumentModel> CachedPassthroughQuery(string filename, string subject)
        //{
        //    List<DocumentModel> documentModels = new List<DocumentModel>();
        //    PassthroughQuery query = new PassthroughQuery();
        //    string queryString = "select r_object_id, object_name,subject from dm_document";

        //    if ((!String.IsNullOrEmpty(filename)) && (!String.IsNullOrEmpty(subject)))
        //    {
        //        queryString = "select r_object_id, object_name,subject from dm_document where object_name like '%" + filename + "%' and subject like '%" + subject + "%'";
        //    }
        //    if ((!String.IsNullOrEmpty(filename)) && (String.IsNullOrEmpty(subject)))
        //    {
        //        queryString = "select r_object_id, object_name,subject from dm_document where object_name like '%" + filename + "%'";
        //    }
        //    if ((String.IsNullOrEmpty(filename)) && (!String.IsNullOrEmpty(subject)))
        //    {
        //        queryString = "select r_object_id, object_name,subject from dm_document where subject like '%" + subject + "%'";
        //    }

        //    query.QueryString = queryString;
        //    //Console.WriteLine("Query string is " + query.QueryString);
        //    query.AddRepository(DefaultRepository);
        //    QueryExecution queryEx = new QueryExecution();
        //    OperationOptions operationOptions = null;

        //    queryEx.CacheStrategyType = CacheStrategyType.BASIC_FILE_CACHE_STRATEGY;
        //    // Console.WriteLine("CacheStrategyType == " + queryEx.CacheStrategyType);
        //    queryEx.MaxResultCount = 10;
        //    //Console.WriteLine("MaxResultCount = " + queryEx.MaxResultCount);
        //    try
        //    {


        //        while (true)
        //        {
        //            QueryResult queryResult = queryService.Execute(query, queryEx,
        //                                                       operationOptions);
        //            DataPackage resultDp = queryResult.DataPackage;

        //            List<DataObject> dataObjects = resultDp.DataObjects;
        //            if (dataObjects.Count == 0)
        //            {
        //                break;
        //            }
        //            //Console.WriteLine("Total objects returned is: " + dataObjects.Count);
        //            foreach (DataObject dObj in dataObjects)
        //            {
        //                PropertySet docProperties = dObj.Properties;
        //                String objectId = dObj.Identity.GetValueAsString();
        //                String docName = docProperties.Get("object_name").GetValueAsString();
        //                string repName = dObj.Identity.RepositoryName;
        //                string docsubject = docProperties.Get("subject").GetValueAsString();
        //                //Console.WriteLine("RepositoryName: " + repName + " ,Document: " + objectId + " ,Name:" + docName + " ,Subject:" + docsubject);
        //                documentModels.Add(new DocumentModel() { ObjectId = objectId, ObjectName = docName, Subject = docsubject });
        //                //yield;
        //            }
        //            queryEx.StartingIndex += 10;
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //        documentModels.Add(new DocumentModel() { ObjectId = "0", ObjectName =  ex.Message, Subject = ex.StackTrace });
        //    }
        //    return documentModels;
        //}

        public IEnumerable<KeywordModel> callQueryServiceKeyword()
        {
            List<KeywordModel> lstKeywords = new List<KeywordModel>();
            PassthroughQuery query = new PassthroughQuery();
            string queryString = "select distinct ecs_pl_category,object_name, ecs_short_val  from ecs_picklist where ecs_pl_category in ('Type of Document', 'Acceptance Code', 'Area', 'Revision', 'Discipline', 'Contract Number', 'Project Reference', 'Issue Reason')";
            //string queryString = "select r_object_id, object_name,subject from dm_document";

            //if ((!String.IsNullOrEmpty(filename)) && (!String.IsNullOrEmpty(subject)))
            //{
            //    queryString = "select r_object_id, object_name,subject from dm_document where object_name like '%" + filename + "%' and subject like '%" + subject + "%'";
            //}
            //if ((!String.IsNullOrEmpty(filename)) && (String.IsNullOrEmpty(subject)))
            //{
            //    queryString = "select r_object_id, object_name,subject from dm_document where object_name like '%" + filename + "%'";
            //}
            //if ((String.IsNullOrEmpty(filename)) && (!String.IsNullOrEmpty(subject)))
            //{
            //    queryString = "select r_object_id, object_name,subject from dm_document where subject like '%" + subject + "%'";
            //}

            query.QueryString = queryString;
            //Console.WriteLine("Query string is " + query.QueryString);
            query.AddRepository(DefaultRepository);
            QueryExecution queryEx = new QueryExecution();
            OperationOptions operationOptions = null;

            queryEx.CacheStrategyType = CacheStrategyType.BASIC_FILE_CACHE_STRATEGY;
            // Console.WriteLine("CacheStrategyType == " + queryEx.CacheStrategyType);
            queryEx.MaxResultCount = 10;
            //Console.WriteLine("MaxResultCount = " + queryEx.MaxResultCount);
            try
            {


                while (true)
                {
                    QueryResult queryResult = queryService.Execute(query, queryEx,
                                                               operationOptions);
                    DataPackage resultDp = queryResult.DataPackage;

                    List<DataObject> dataObjects = resultDp.DataObjects;
                    if (dataObjects.Count == 0)
                    {
                        break;
                    }
                    //Console.WriteLine("Total objects returned is: " + dataObjects.Count);
                    foreach (DataObject dObj in dataObjects)
                    {
                        PropertySet docProperties = dObj.Properties;
                        String category = docProperties.Get("ecs_pl_category").GetValueAsString();
                        String ObjectName = docProperties.Get("object_name").GetValueAsString();
                        String SV = docProperties.Get("ecs_short_val").GetValueAsString();

                        lstKeywords.Add(new KeywordModel() { Category = category, ObjectName = ObjectName, ShortValue = SV });
                        //yield;
                    }
                    queryEx.StartingIndex += 10;
                }
            }
            catch (Exception ex)
            {
                lstKeywords.Add(new KeywordModel() { Category = "Failure", ObjectName = "Check", ShortValue = "CK" });
                lstKeywords.Add(new KeywordModel() { Category = "Exception", ObjectName = ex.StackTrace, ShortValue = ex.Message });
            }
            return lstKeywords;
        }
     
    }
}
