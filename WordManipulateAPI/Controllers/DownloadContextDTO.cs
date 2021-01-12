using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WordManipulateAPI.Controllers
{
    public class DownloadContextDTO
    {
        public string object_id { get; set; }
        public int RidType;
        public string category;
        public string SaveFilename;
        public int attachsequence;
        public string attachmentguid;
    }
}