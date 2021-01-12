using DocumentFormat.OpenXml;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WordManipulateAPI.Controllers;

namespace WordManipulateAPI.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Home Page - testing";

            //EPFMController oEPFMController = new EPFMController();

            //DownloadContextDTO oDownloadContextDTO = new DownloadContextDTO();
            //oDownloadContextDTO.object_id = "09005ba080c90e52";

            //oDownloadContextDTO.attachmentguid
            //oEPFMController.DownloadEPFMDocument(oDownloadContextDTO);
            return View();
        }


    }
}
