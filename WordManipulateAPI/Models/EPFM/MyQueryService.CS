
using System;
using System.Collections.Generic;

using Emc.Documentum.FS.Services.Core;
using Emc.Documentum.FS.DataModel.Core;
using Emc.Documentum.FS.DataModel.Core.Query;
using Emc.Documentum.FS.DataModel.Core.Properties;
using Emc.Documentum.FS.Runtime.Context;
using Emc.Documentum.FS.DataModel.Core.Context;

namespace WordManipulateAPI.Models.EPFM
{

    public class MyQueryService
    {
        /************************************************************
         * You must supply valid values for the following fields: */

        /* The repository that you want to run the query on */
        private String repository = "er_dev_epfm_01";

        /* The username to login to the repository */
        private String userName = "svc.cmsadmindev";

        /* The password for the username */
        private String password = "Fc$YqK96d8%C";

        /* The address where the DFS services are located */
        private String address = "http://er-kdc-ddcpas01:8082/dfs/services";

        /***********************************************************/

        /* The module name for the DFS core services */
        private static String moduleName = "core";

        private IServiceContext serviceContext;

        

        public void setContext(string userName, string password, string address, string repository)
        {
            /* 
             * Get the service context and set the user 
             * credentials and repository information
             */
            this.userName = userName;
            this.password = password;
            this.repository = repository;
            this.address = address;

            ContextFactory contextFactory = ContextFactory.Instance;
            serviceContext = contextFactory.NewContext();
            RepositoryIdentity repositoryIdentity =
                new RepositoryIdentity(repository, userName, password, "");
            serviceContext.AddIdentity(repositoryIdentity);
        }

        /*
         * Demonstrates a typical scenario for calling the query service.
         * Gets an instance of the Query service and calls the execute operation
         * with a hard-coded query and operation options.
         */
        public List<CabinetModel> callQueryServiceCabinet()
        {
            /*
             * Get an instance of the QueryService by passing 
             * in the service context to the service factory.
             */
            string result = "start";
            List<CabinetModel> lstCabinets = new List<CabinetModel>();
            try
            {


                ServiceFactory serviceFactory = ServiceFactory.Instance;
                IQueryService querySvc = serviceFactory.GetRemoteService<IQueryService>(serviceContext, moduleName, address);

                /*
                * Construct the query and the QueryExecution options
                */
                PassthroughQuery query = new PassthroughQuery();
                query.QueryString = "select r_object_id, object_name from dm_cabinet";
                query.AddRepository(repository);
                QueryExecution queryEx = new QueryExecution();
                queryEx.CacheStrategyType = CacheStrategyType.DEFAULT_CACHE_STRATEGY;

                /*
                 * Execute the query passing in the operation options and print the result
                 */
                OperationOptions operationOptions = null;
                QueryResult queryResult = querySvc.Execute(query, queryEx, operationOptions);
                result = result + Environment.NewLine + "QueryId == " + query.QueryString;
                result = result + Environment.NewLine + "CacheStrategyType == " + queryEx.CacheStrategyType;
                //Console.WriteLine("QueryId == " + query.QueryString);
                //Console.WriteLine("CacheStrategyType == " + queryEx.CacheStrategyType);
                DataPackage resultDp = queryResult.DataPackage;
                List<DataObject> dataObjects = resultDp.DataObjects;
                int numberOfObjects = dataObjects.Count;
                result = result + Environment.NewLine + "Total objects returned is: " + numberOfObjects;
                //Console.WriteLine("Total objects returned is: " + numberOfObjects);
                foreach (DataObject dObj in dataObjects)
                {
                    PropertySet docProperties = dObj.Properties;
                    String objectId = dObj.Identity.GetValueAsString();
                    String docName = docProperties.Get("object_name").GetValueAsString();
                    //Console.WriteLine("Document " + objectId + " name is " + docName);
                    result = result + Environment.NewLine + "Document " + objectId + " name is " + docName;
                    lstCabinets.Add(new CabinetModel() { ObjectId = objectId, ObjectName = docName });
                }
            }
            //catch (FaultException<SerializableException> ex)
            //{
            //    Console.WriteLine(String.Format("Got FaultException[{0}] with message: {1}\n", ex.Detail, ex.Message));
            //}
            catch (Exception exx)
            {
                // Console.WriteLine(exx.StackTrace);
                throw exx;
            }
            return lstCabinets;
        }

        public List<CabinetModel> callQueryServiceDocument(string cabinetid)
        {
            /*
             * Get an instance of the QueryService by passing 
             * in the service context to the service factory.
             */
            string result = "start";
            List<CabinetModel> lstCabinets = new List<CabinetModel>();
            try
            {


                ServiceFactory serviceFactory = ServiceFactory.Instance;
                IQueryService querySvc = serviceFactory.GetRemoteService<IQueryService>(serviceContext, moduleName, address);

                /*
                * Construct the query and the QueryExecution options
                */
                PassthroughQuery query = new PassthroughQuery();
                query.QueryString = "select r_object_id, object_name from dm_document";
                query.AddRepository(repository);
                QueryExecution queryEx = new QueryExecution();
                queryEx.CacheStrategyType = CacheStrategyType.DEFAULT_CACHE_STRATEGY;

                /*
                 * Execute the query passing in the operation options and print the result
                 */
                OperationOptions operationOptions = null;
                QueryResult queryResult = querySvc.Execute(query, queryEx, operationOptions);
                result = result + Environment.NewLine + "QueryId == " + query.QueryString;
                result = result + Environment.NewLine + "CacheStrategyType == " + queryEx.CacheStrategyType;
                //Console.WriteLine("QueryId == " + query.QueryString);
                //Console.WriteLine("CacheStrategyType == " + queryEx.CacheStrategyType);
                DataPackage resultDp = queryResult.DataPackage;
                List<DataObject> dataObjects = resultDp.DataObjects;
                int numberOfObjects = dataObjects.Count;
                result = result + Environment.NewLine + "Total objects returned is: " + numberOfObjects;
                //Console.WriteLine("Total objects returned is: " + numberOfObjects);
                foreach (DataObject dObj in dataObjects)
                {
                    PropertySet docProperties = dObj.Properties;
                    String objectId = dObj.Identity.GetValueAsString();
                    String docName = docProperties.Get("object_name").GetValueAsString();
                    //Console.WriteLine("Document " + objectId + " name is " + docName);
                    result = result + Environment.NewLine + "Document " + objectId + " name is " + docName;
                    lstCabinets.Add(new CabinetModel() { ObjectId = objectId, ObjectName = docName });
                }
            }
            //catch (FaultException<SerializableException> ex)
            //{
            //    Console.WriteLine(String.Format("Got FaultException[{0}] with message: {1}\n", ex.Detail, ex.Message));
            //}
            catch (Exception exx)
            {
                // Console.WriteLine(exx.StackTrace);
                throw exx;
            }
            return lstCabinets;
        }

        public List<KeywordModel> callQueryServiceKeyword()
        {
            /*
             * Get an instance of the QueryService by passing 
             * in the service context to the service factory.
             */
            string result = "start";
            List<KeywordModel> lstKeywords = new List<KeywordModel>();
            try
            {


                ServiceFactory serviceFactory = ServiceFactory.Instance;
                IQueryService querySvc = serviceFactory.GetRemoteService<IQueryService>(serviceContext, moduleName, address);

                /*
                * Construct the query and the QueryExecution options
                */
                PassthroughQuery query = new PassthroughQuery();
                query.QueryString = "select *  from ecs_picklist where ecs_pl_category in ('Type of Document', 'Acceptance Code', 'Area', 'Revision', 'Discipline', 'Contract Number', 'Project Reference', 'Issue Reason')";
                query.AddRepository(repository);
                QueryExecution queryEx = new QueryExecution();
                queryEx.CacheStrategyType = CacheStrategyType.NO_CACHE_STRATEGY;

                /*
                 * Execute the query passing in the operation options and print the result
                 */
                OperationOptions operationOptions = null;
                QueryResult queryResult = querySvc.Execute(query, queryEx, operationOptions);
                result = result + Environment.NewLine + "QueryId == " + query.QueryString;
                result = result + Environment.NewLine + "CacheStrategyType == " + queryEx.CacheStrategyType;
                //Console.WriteLine("QueryId == " + query.QueryString);
                //Console.WriteLine("CacheStrategyType == " + queryEx.CacheStrategyType);
                DataPackage resultDp = queryResult.DataPackage;
                List<DataObject> dataObjects = resultDp.DataObjects;
                int numberOfObjects = dataObjects.Count;
                result = result + Environment.NewLine + "Total objects returned is: " + numberOfObjects;
                //Console.WriteLine("Total objects returned is: " + numberOfObjects);
                foreach (DataObject dObj in dataObjects)
                {
                    PropertySet docProperties = dObj.Properties;
                    String category = docProperties.Get("ecs_pl_category").GetValueAsString();
                    String ObjectName = docProperties.Get("object_name").GetValueAsString();
                    String SV = docProperties.Get("ecs_short_val").GetValueAsString();

                    lstKeywords.Add(new KeywordModel() { Category = category, ObjectName = ObjectName, ShortValue = SV });
                }
            }
            //catch (FaultException<SerializableException> ex)
            //{
            //    Console.WriteLine(String.Format("Got FaultException[{0}] with message: {1}\n", ex.Detail, ex.Message));
            //}
            catch (Exception )
            {
                // Console.WriteLine(exx.StackTrace);
                throw ;
            }
            return lstKeywords;
        }

    }
}
