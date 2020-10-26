using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Mvc;
using WordManipulateAPI.Helpers;

namespace WordManipulateAPI.Controllers
{
    public class OCRController : ApiController
    {
        [System.Web.Http.HttpGet]
        public string SaveFileOCRFormat()
        {
            OCRHelper ocrHelper = new OCRHelper();
            ocrHelper.OCR(false);
            return "abc";            
        }

       
    }
}
