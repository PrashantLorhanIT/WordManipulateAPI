using Emc.Documentum.FS.Services.Core;
using Emc.Documentum.FS.DataModel.Core;
using Emc.Documentum.FS.DataModel.Core.Query;
using Emc.Documentum.FS.Services.Search;
using Emc.Documentum.FS.Runtime;
using Emc.Documentum.FS.Runtime.Context;
using Emc.Documentum.FS;
using System;
using System.Collections.Generic;
using System.Text;
using Emc.Documentum.FS.DataModel.Core.Profiles;
using Emc.Documentum.FS.DataModel.Core.Properties;

namespace WordManipulateAPI.Models.EPFM
{
    public class SearchServiceDemo: DemoBase
    {
        private ISearchService searchService;
        
        public SearchServiceDemo(String defaultRepository, String secondaryRepository, String userName, String password, String domain)
            : base(defaultRepository, secondaryRepository, userName, password)
        {
            ServiceFactory serviceFactory = ServiceFactory.Instance;
            searchService =
                serviceFactory.GetRemoteService<ISearchService>(DemoServiceContext, "core", domain);
        }

        public List<Repository> RepositoryList()
        {
            List<Repository> repositoryList = searchService.GetRepositoryList(null);
            foreach (Repository r in repositoryList)
            {
                Console.WriteLine(r.Name);
                Console.Write("-" + r.Properties.UserLoginCapability);
            }
            return repositoryList;
        }

        public QueryResult SimplePassthroughQueryCabinet()
        {
            QueryResult queryResult;
            string queryString = "select r_object_id, object_name from dm_cabinet";
            int startingIndex = 0;
            int maxResults = 60;
            int maxResultsPerSource = 20;

            PassthroughQuery q = new PassthroughQuery();
            q.QueryString = queryString;
            q.AddRepository(DefaultRepository);

            QueryExecution queryExec = new QueryExecution(startingIndex,
                                                          maxResults,
                                                          maxResultsPerSource);
            queryExec.CacheStrategyType = CacheStrategyType.NO_CACHE_STRATEGY;

            queryResult = searchService.Execute(q, queryExec, null);

            QueryStatus queryStatus = queryResult.QueryStatus;
            RepositoryStatusInfo repStatusInfo = queryStatus.RepositoryStatusInfos[0];
            if (repStatusInfo.Status == Status.FAILURE)
            {
                Console.WriteLine(repStatusInfo.ErrorTrace);
                throw new Exception("Query failed to return result.");
            }
            Console.WriteLine("Query returned result successfully.");
            DataPackage dp = queryResult.DataPackage;
            Console.WriteLine("DataPackage contains " + dp.DataObjects.Count + " objects.");
            foreach (DataObject dObj in dp.DataObjects)
            {
                //Console.WriteLine(dObj.Identity.GetValueAsString());
                //
                PropertySet docProperties = dObj.Properties;
                String objectId = dObj.Identity.GetValueAsString();
                String docName = docProperties.Get("object_name").GetValueAsString();
                
                Console.WriteLine("Document " + objectId + " name is " + docName);
               // result = result + Environment.NewLine + "Document " + objectId + " name is " + docName;
                //lstCabinets.Add(new CabinetModel() { ObjectId = objectId, ObjectName = docName });
                //

            }
            return queryResult;
        }
        public QueryResult SimplePassthroughQueryDocument()
        {
            QueryResult queryResult;
            string queryString = "select r_object_id, object_name from eifx_deliverable_doc";
            int startingIndex = 0;
            int maxResults = 60;
            int maxResultsPerSource = 30;

            PassthroughQuery q = new PassthroughQuery();
            q.QueryString = queryString;
            q.AddRepository(DefaultRepository);

            QueryExecution queryExec = new QueryExecution(startingIndex,
                                                          maxResults,
                                                          maxResultsPerSource);
            queryExec.CacheStrategyType = CacheStrategyType.NO_CACHE_STRATEGY;

            queryResult = searchService.Execute(q, queryExec, null);

            QueryStatus queryStatus = queryResult.QueryStatus;
            RepositoryStatusInfo repStatusInfo = queryStatus.RepositoryStatusInfos[0];
            if (repStatusInfo.Status == Status.FAILURE)
            {
                Console.WriteLine(repStatusInfo.ErrorTrace);
                throw new Exception("Query failed to return result.");
            }
            Console.WriteLine("Query returned result successfully.");
            DataPackage dp = queryResult.DataPackage;
            Console.WriteLine("DataPackage contains " + dp.DataObjects.Count + " objects.");
            foreach (DataObject dObj in dp.DataObjects)
            {
                //Console.WriteLine(dObj.Identity.GetValueAsString());
                //
                PropertySet docProperties = dObj.Properties;
                String objectId = dObj.Identity.GetValueAsString();
                String docName = docProperties.Get("object_name").GetValueAsString();

                Console.WriteLine("Document " + objectId + " name is " + docName);
                // result = result + Environment.NewLine + "Document " + objectId + " name is " + docName;
                //lstCabinets.Add(new CabinetModel() { ObjectId = objectId, ObjectName = docName });
                //

            }
            return queryResult;
        }

        public IEnumerable<DocumentModel> SimplePassthroughQueryDocumentWithPath_New(SearchModel search)
        {

            List<DocumentModel> documentModels = new List<DocumentModel>();
            QueryResult queryResult;
            try
            {
                //string queryString = "select dm_document.r_object_id, dm_document.subject, dm_document.a_content_type,dm_document.object_name,dm_format.dos_extension from dm_document,dm_format where  dm_document.a_content_type = dm_format.name";
                string queryString = " select r_object_id ,title ,object_name,r_creation_date"
+ ", r_modify_date, eif_date_due, eif_acceptance_code, eif_area, eif_revision"
+ ", eif_issue_reason, eif_discipline, eif_originator, eif_type_of_doc, eif_project_ref"
+ ", er_actual_sub_date, er_wbs_level3, er_wbs_level4, er_package_name, er_contract_number"
+ ", dm_document.subject, dm_document.a_content_type, dm_format.dos_extension from eifx_deliverable_doc,dm_document,dm_format where eifx_deliverable_doc.r_object_id = dm_document.r_object_id and dm_document.a_content_type = dm_format.name ";

//                string queryString = " select eifx_deliverable_doc.r_object_id ,eifx_deliverable_doc.title ,eifx_deliverable_doc.object_name,eifx_deliverable_doc.r_creation_date"
//+ ",eifx_deliverable_doc.r_modify_date,eifx_deliverable_doc.eif_date_due,eifx_deliverable_doc.eif_acceptance_code,eifx_deliverable_doc.eif_area,eifx_deliverable_doc.eif_revision"
//+ ",eifx_deliverable_doc.eif_issue_reason,eifx_deliverable_doc.eif_discipline,eifx_deliverable_doc.eif_originator,eifx_deliverable_doc.eif_type_of_doc,eifx_deliverable_doc.eif_project_ref"
//+ ",eifx_deliverable_doc.er_actual_sub_date,eifx_deliverable_doc.er_wbs_level3,eifx_deliverable_doc.er_wbs_level4,eifx_deliverable_doc.er_package_name,eifx_deliverable_doc.er_contract_number"
//+ ",dm_document.subject,dm_document.a_content_type,dm_format.dos_extensionfrom eifx_deliverable_doc,dm_document,dm_format ";
                //where eifx_deliverable_doc.r_object_id = dm_document.r_object_id and dm_document.a_content_type = dm_format.name 
                if (!String.IsNullOrEmpty(search.ContractNumber))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.er_contract_number like '%" + search.ContractNumber + "%' ";
                }

                if (!String.IsNullOrEmpty(search.DocumentNumber))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.object_name like '%" + search.DocumentNumber + "%' ";
                }

                if (!String.IsNullOrEmpty(search.DocumentTitle))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.title like '%" + search.DocumentTitle + "%' ";
                }

                if (!String.IsNullOrEmpty(search.PackageName))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.er_package_name like '%" + search.PackageName + "%' ";
                }

                if (!String.IsNullOrEmpty(search.DocumentType))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.eif_type_of_doc like '%" + search.DocumentType + "%' ";
                }

                if (!String.IsNullOrEmpty(search.DocumentAcceptanceStatus))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.eif_acceptance_code like '%" + search.DocumentAcceptanceStatus + "%' ";
                }

                if (!String.IsNullOrEmpty(search.Originator))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.eif_originator like '%" + search.Originator + "%' ";
                }

                if (!String.IsNullOrEmpty(search.Area))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.eif_area like '%" + search.Area + "%' ";
                }

                if (!String.IsNullOrEmpty(search.Revision))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.eif_revision like '%" + search.Revision + "%' ";
                }

                if (!String.IsNullOrEmpty(search.Discipline))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.eif_discipline like '%" + search.Discipline + "%' ";
                }

                //if (!String.IsNullOrEmpty(search.Filename))
                //{
                //    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " er_contract_number like '%" + search.Filename + "%' ";
                //}

                if (!String.IsNullOrEmpty(search.ProjectReference))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.eif_project_ref like '%" + search.ProjectReference + "%' ";
                }

                if (!String.IsNullOrEmpty(search.IssueReason))
                {
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " eifx_deliverable_doc.eif_issue_reason like '%" + search.IssueReason + "%' ";
                }

                //if (!String.IsNullOrEmpty(search.SuperSearch))
                //{
                //    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " er_contract_number like '" + search.SuperSearch + "' ";
                //}

                //if (!String.IsNullOrEmpty(search.Location))
                //{
                //    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + " er_contract_number like '" + search.Location + "' ";
                //}

                if ( (!String.IsNullOrEmpty(search.DateRangeDay)) && (!String.IsNullOrEmpty(search.DateRangeType)))
                {
                    string type = "";
                    if (search.DateRangeType == "Modified Date")
                    {
                        type = "r_modify_date";
                    }
                    else if (search.DateRangeType == "Created Date")
                    {
                        type = "r_creation_date";
                    }
                    else if (search.DateRangeType == "yyMSG_Date_Submittedyy")
                    {
                        type = " er_actual_sub_date";
                    }
                    else if (search.DateRangeType == "yyMSG_DueDateyy")
                    {
                        type = "eif_date_due";
                    }


                    if (search.DateRangeDay == "Today")
                    {
                        type = type + " between " + DateTime.Today.ToString("MM/DD/yyyy 00:00:00") + " and " + DateTime.Today.ToString("MM/DD/yyyy 23:59:00");
                    }
                    else if(search.DateRangeDay == "Yesterday")
                    {
                        type = type + " between " + DateTime.Today.AddDays(-1).ToString("MM/DD/yyyy 00:00:00") + " and " + DateTime.Today.AddDays(-1).ToString("MM/DD/yyyy 23:59:00");
                    }
                    else if (search.DateRangeDay == "Last 7 days")
                    {
                        type = type + " between " + DateTime.Today.AddDays(-7).ToString("MM/DD/yyyy 00:00:00") + " and " + DateTime.Today.AddDays(0).ToString("MM/DD/yyyy 23:59:00");
                    }
                    else if (search.DateRangeDay == "Last 30 days")
                    {
                        type = type + " between " + DateTime.Today.AddDays(-30).ToString("MM/DD/yyyy 00:00:00") + " and " + DateTime.Today.AddDays(0).ToString("MM/DD/yyyy 23:59:00");
                    }
                    else if (search.DateRangeDay == "Last 90 days")
                    {
                        type = type + " between " + DateTime.Today.AddDays(-90).ToString("MM/DD/yyyy 00:00:00") + " and " + DateTime.Today.AddDays(0).ToString("MM/DD/yyyy 23:59:00");
                    }
                    else if (search.DateRangeDay == "Between")
                    {
                        type = type + " between " + search.DateRangeFrom.ToString("MM/DD/yyyy 00:00:00") + " and " + search.DateRangeTo.ToString("MM/DD/yyyy 23:59:00");
                    }
                    else if (search.DateRangeDay == "Before")
                    {
                        type = type + " <= " + search.DateRangeTo.ToString("MM/DD/yyyy 23:59:00");
                    }
                    else if (search.DateRangeDay == "After")
                    {
                        type = type + " >= " + search.DateRangeFrom.ToString("MM/DD/yyyy 00:00:00");
                    }
                    else if (search.DateRangeDay == "On")
                    {
                        type = type + " between " + search.DateRangeFrom.ToString("MM/DD/yyyy 00:00:00") + " and " + search.DateRangeFrom.ToString("MM/DD/yyyy 23:59:00");
                    }
                    queryString = queryString + (queryString.Contains("where") ? " and " : " where ") + type;
                }




                int startingIndex = 0;
                int maxResults = 60;
                int maxResultsPerSource = 20;

                PassthroughQuery q = new PassthroughQuery();
                q.QueryString = queryString;
                q.AddRepository(DefaultRepository);

                QueryExecution queryExec = new QueryExecution(startingIndex,
                                                              maxResults,
                                                              maxResultsPerSource);
                queryExec.CacheStrategyType = CacheStrategyType.NO_CACHE_STRATEGY;

                queryResult = searchService.Execute(q, queryExec, null);

                QueryStatus queryStatus = queryResult.QueryStatus;
                RepositoryStatusInfo repStatusInfo = queryStatus.RepositoryStatusInfos[0];
                if (repStatusInfo.Status == Status.FAILURE)
                {
                    //  Console.WriteLine(repStatusInfo.ErrorTrace);
                    documentModels.Add(new DocumentModel() { ObjectId = "0", ObjectName = repStatusInfo.ErrorMessage, DocumentTitle = repStatusInfo.Name, DocumentNumber = repStatusInfo.Name });
                }
                //Console.WriteLine("Query returned result successfully.");
                DataPackage dp = queryResult.DataPackage;
                //Console.WriteLine("DataPackage contains " + dp.DataObjects.Count + " objects.");
                foreach (DataObject dObj in dp.DataObjects)
                {
                    PropertySet docProperties = dObj.Properties;
                    String objectId = dObj.Identity.GetValueAsString();
                    String docName = docProperties.Get("object_name").GetValueAsString();
                    String Extension = docProperties.Get("dos_extension").GetValueAsString();
                    string repName = dObj.Identity.RepositoryName;
                    string doctitle = docProperties.Get("title").GetValueAsString();
                    string docNumber = docProperties.Get("object_name").GetValueAsString();
                    //Console.WriteLine("RepositoryName: " + repName + " ,Document: " + objectId + " ,Name:" + docName + " ,Subject:" + docsubject);
                    string revision = docProperties.Get("eif_revision").GetValueAsString();
                    string creationDate = docProperties.Get("r_creation_date").GetValueAsString();
                    string package = docProperties.Get("er_package_name").GetValueAsString();
                    string issueReason = docProperties.Get("eif_issue_reason").GetValueAsString();
                    string contract = docProperties.Get("er_contract_number").GetValueAsString();
                    string originator = docProperties.Get("eif_originator").GetValueAsString();
                    string discipline = docProperties.Get("eif_discipline").GetValueAsString();
                    string acceptanceCode = docProperties.Get("eif_acceptance_code").GetValueAsString();
                    string actualSubDate = docProperties.Get("er_actual_sub_date").GetValueAsString();

                    documentModels.Add(new DocumentModel() { ObjectId = objectId, ObjectName = docName + "." + Extension, DocumentTitle = doctitle, DocumentNumber = docNumber,
                                                            Revision = revision, AcceptanceCode = acceptanceCode, ActualSubDate = actualSubDate, Contract = contract,
                                                            CreationDate = creationDate, Discipline = discipline, IssueReason = issueReason, Originator = originator, Package = package
                                                            });

                }
            }
            catch (Exception ex)
            {

                documentModels.Add(new DocumentModel() { ObjectId = "1", ObjectName = ex.Message, DocumentTitle = ex.StackTrace , DocumentNumber = "XXXX-1FD-345-QWE",
                                                            Revision = "Rev983475",
                                                            AcceptanceCode = "TRRTT",
                                                            ActualSubDate = "21-09-2020",
                                                            Contract = "P34242-Contract",
                                                            CreationDate = "21-09-2020",
                                                            Discipline = "FR8530498",
                                                            IssueReason = "Approved",
                                                            Originator = "Mark Boyle",
                                                            Package = "DFGH54938"
                                                        });
            }

            return documentModels;
        }

        //public IEnumerable<DocumentModel> SimplePassthroughQueryDocumentWithPath(string filename, string subject)
        //{

        //    List<DocumentModel> documentModels = new List<DocumentModel>();
        //    QueryResult queryResult;
        //    try
        //    {
        //        string queryString = "select dm_document.r_object_id, dm_document.subject, dm_document.a_content_type,dm_document.object_name,dm_format.dos_extension from dm_document,dm_format where  dm_document.a_content_type = dm_format.name";

        //        if ((!String.IsNullOrEmpty(filename)) && (!String.IsNullOrEmpty(subject)))
        //        {
        //            queryString = queryString + " and upper(object_name) like '%" + filename.ToUpper() + "%' and upper(subject) like '%" + subject.ToUpper() + "%'";
        //        }
        //        if ((!String.IsNullOrEmpty(filename)) && (String.IsNullOrEmpty(subject)))
        //        {
        //            queryString = queryString + " and upper(object_name) like '%" + filename.ToUpper() + "%'";
        //        }
        //        if ((String.IsNullOrEmpty(filename)) && (!String.IsNullOrEmpty(subject)))
        //        {
        //            queryString = queryString + " and upper(subject) like '%" + subject.ToUpper() + "%'";
        //        }
        //        int startingIndex = 0;
        //        int maxResults = 60;
        //        int maxResultsPerSource = 20;

        //        PassthroughQuery q = new PassthroughQuery();
        //        q.QueryString = queryString;
        //        q.AddRepository(DefaultRepository);

        //        QueryExecution queryExec = new QueryExecution(startingIndex,
        //                                                      maxResults,
        //                                                      maxResultsPerSource);
        //        queryExec.CacheStrategyType = CacheStrategyType.NO_CACHE_STRATEGY;

        //        queryResult = searchService.Execute(q, queryExec, null);

        //        QueryStatus queryStatus = queryResult.QueryStatus;
        //        RepositoryStatusInfo repStatusInfo = queryStatus.RepositoryStatusInfos[0];
        //        if (repStatusInfo.Status == Status.FAILURE)
        //        {
        //            //  Console.WriteLine(repStatusInfo.ErrorTrace);
        //            documentModels.Add(new DocumentModel() { ObjectId = "0", ObjectName = repStatusInfo.ErrorMessage, Subject = repStatusInfo.ErrorTrace });
        //        }
        //        //Console.WriteLine("Query returned result successfully.");
        //        DataPackage dp = queryResult.DataPackage;
        //        //Console.WriteLine("DataPackage contains " + dp.DataObjects.Count + " objects.");
        //        foreach (DataObject dObj in dp.DataObjects)
        //        {
        //            PropertySet docProperties = dObj.Properties;
        //            String objectId = dObj.Identity.GetValueAsString();
        //            String docName = docProperties.Get("object_name").GetValueAsString();
        //            String Extension = docProperties.Get("dos_extension").GetValueAsString();
        //            string repName = dObj.Identity.RepositoryName;
        //            string docsubject = docProperties.Get("subject").GetValueAsString();
        //            //Console.WriteLine("RepositoryName: " + repName + " ,Document: " + objectId + " ,Name:" + docName + " ,Subject:" + docsubject);

        //            documentModels.Add(new DocumentModel() { ObjectId = objectId, ObjectName = docName+"."+Extension , Subject = docsubject });

        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //        documentModels.Add(new DocumentModel() { ObjectId = "1", ObjectName = "SampleFile.txt", Subject = "This is a Sample" });
        //    }
            
        //    return documentModels;
        //}
        public void SimpleStructuredQuery(String docName)
        {
            String repoName = DefaultRepository;
            Console.WriteLine("Called SimpleStructuredQuery - " + DefaultRepository);
            PropertyProfile propertyProfile = new PropertyProfile();
            propertyProfile.FilterMode = PropertyFilterMode.IMPLIED;
            OperationOptions operationOptions = new OperationOptions();
            operationOptions.Profiles.Add(propertyProfile);

            // Create query
            StructuredQuery q = new StructuredQuery();
            q.AddRepository(repoName);
            q.ObjectType = "dm_document";
            q.IsIncludeHidden = true;
            q.IsDatabaseSearch = true;
            ExpressionSet expressionSet = new ExpressionSet();
            expressionSet.AddExpression(new PropertyExpression("object_name",
                                                               Condition.CONTAINS,
                                                               docName));
            q.RootExpressionSet = expressionSet;

            // Execute Query
            int startingIndex = 0;
            int maxResults = 60;
            int maxResultsPerSource = 20;
            QueryExecution queryExec = new QueryExecution(startingIndex,
                                                          maxResults,
                                                          maxResultsPerSource);
            QueryResult queryResult = searchService.Execute(q, queryExec, operationOptions);

            QueryStatus queryStatus = queryResult.QueryStatus;
            RepositoryStatusInfo repStatusInfo = queryStatus.RepositoryStatusInfos[0];
            if (repStatusInfo.Status == Status.FAILURE)
            {
                Console.WriteLine(repStatusInfo.ErrorTrace);
                throw new Exception("Query failed to return result.");
            }
            Console.WriteLine("Query returned result successfully.");

            // print results
            Console.WriteLine("DataPackage contains " + queryResult.DataObjects.Count + " objects.");
            foreach (DataObject dataObject in queryResult.DataObjects)
            {
                Console.WriteLine(dataObject.Identity.GetValueAsString());
            }
        }

        //public IEnumerable<KeywordModel> SimplePassthroughQueryForKeywords()
        //{

        //    List<KeywordModel> keywordModels = new List<KeywordModel>();
        //    QueryResult queryResult;
        //    try
        //    {
        //        string queryString = "select distinct ecs_pl_category,object_name, ecs_short_val  from ecs_picklist where ecs_pl_category in ('Type of Document', 'Acceptance Code', 'Area', 'Revision', 'Discipline', 'Contract Number', 'Project Reference', 'Issue Reason')";
        //        int startingIndex = 0;
        //        int maxResults = 60;
        //        int maxResultsPerSource = 20;

        //        PassthroughQuery q = new PassthroughQuery();
        //        q.QueryString = queryString;
        //        q.AddRepository(DefaultRepository);

        //        QueryExecution queryExec = new QueryExecution(startingIndex,
        //                                                      maxResults,
        //                                                      maxResultsPerSource);
        //        queryExec.CacheStrategyType = CacheStrategyType.NO_CACHE_STRATEGY;

        //        queryResult = searchService.Execute(q, queryExec, null);

        //        QueryStatus queryStatus = queryResult.QueryStatus;
        //        RepositoryStatusInfo repStatusInfo = queryStatus.RepositoryStatusInfos[0];
        //        if (repStatusInfo.Status == Status.FAILURE)
        //        {
        //            //  Console.WriteLine(repStatusInfo.ErrorTrace);
        //            keywordModels.Add(new KeywordModel() { Category = "Failure",  ObjectName = "Check", ShortValue = "CK" });
        //        }
        //        //Console.WriteLine("Query returned result successfully.");
        //        DataPackage dp = queryResult.DataPackage;
        //        //Console.WriteLine("DataPackage contains " + dp.DataObjects.Count + " objects.");
        //        foreach (DataObject dObj in dp.DataObjects)
        //        {
        //            PropertySet docProperties = dObj.Properties;
        //            String category = docProperties.Get("ecs_pl_category").GetValueAsString();
        //            String ObjectName = docProperties.Get("object_name").GetValueAsString();
        //            String SV = docProperties.Get("ecs_short_val").GetValueAsString();

        //            keywordModels.Add(new KeywordModel() { Category = category, ObjectName =   ObjectName, ShortValue = SV });

        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //        keywordModels.Add(new KeywordModel() { Category = "Exception", ObjectName = ex.StackTrace, ShortValue = ex.Message });
        //    }

        //    return keywordModels;
        //}
    }

}
