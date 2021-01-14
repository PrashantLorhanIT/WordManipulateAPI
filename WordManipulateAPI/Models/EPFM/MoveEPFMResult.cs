using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WordManipulateAPI.Models.EPFM
{
    public class MoveEPFMResult
    {
        public int RidAttachment { get; set; }
        public string Filename { get; set; }
        public string PreviewPath { get; set; }
    }

    public class MoveEPFMResultWrapper
    {
        public MoveEPFMResult data { get; set; }
        public string message { get; set; }
        public string statusCode { get; set; }
    }

    public class Usermaster
    {
        public Usermaster data { get; set; }
        public string EpfmUsername { get; set; }
        public string EpfmPassword { get; set; }
        public string Isaduser { get; set; }

    }
}