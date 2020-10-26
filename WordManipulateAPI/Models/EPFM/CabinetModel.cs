using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WordManipulateAPI.Models.EPFM
{
    public class CabinetModel
    {
        public string ObjectId { get; set; }
        public string ObjectName { get; set; }
    }
    public class DocumentModel
    {
        public string ObjectId { get; set; }
        public string ObjectName { get; set; }
        public string DocumentNumber { get; set; }
        public string DocumentTitle { get; set; }
    }

    public class KeywordModel
    {
        public string ObjectName { get; set; }
        public string Category { get; set; }
        public string ShortValue { get; set; }
    }

    public class SearchModel
    {
        
        public string Filename { get; set; }

        public string DocumentNumber { get; set; }

        public string DocumentTitle { get; set; }

        public string PackageName { get; set; }

        public string DocumentType { get; set; }

        public string DocumentAcceptanceStatus { get; set; }

        public string Originator { get; set; }

        public string Area { get; set; }

        public string Revision { get; set; }

        public string Discipline { get; set; }
        public string ContractNumber { get; set; }

        public string ProjectReference { get; set; }


        public string IssueReason { get; set; }

        public string DateRangeType { get; set; }

        public string DateRangeDay { get; set; }

        public DateTime DateRangeFrom { get; set; }

        public DateTime DateRangeTo { get; set; }

        public string SuperSearch { get; set; }
    }
}