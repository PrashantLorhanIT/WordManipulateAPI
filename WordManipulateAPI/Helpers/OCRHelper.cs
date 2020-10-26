using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml.ExtendedProperties;
using Dynamsoft.Core;
using Dynamsoft.OCR;
using Dynamsoft.OCR.Enums;
using Dynamsoft.PDF;

namespace WordManipulateAPI.Helpers
{

    public class OCRHelper
    {
        private string m_strCurrentDirectory;
        private string m_StrProductKey = "t0068MgAAAJfb3IyyncwbZ90LOgaKFZgxkXldgAG5cFF4QIP5NYoLFM1mF5ZU01A1EX8QD6MUtNkZycxCYs42WEgocYOwGRk=";
        private ImageCore m_ImageCore = null;
        private Tesseract m_Tesseract = null;
        private PDFRasterizer m_PDFRasterizer = null;
        Dictionary<string, string> languages = new Dictionary<string, string>();
        string[] resultFormat = new string[] { "Text File", "Adobe PDF Plain Text File", "Adobe PDF Image Over Text File" };
        List<Bitmap> tempListSelectedBitmap = null;
        public OCRHelper()
        {
            languages.Add("English", "eng");
            m_strCurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            m_Tesseract = new Tesseract(m_StrProductKey);
            m_ImageCore = new ImageCore();
            m_PDFRasterizer = new PDFRasterizer(m_StrProductKey);
            tempListSelectedBitmap = new List<Bitmap>();
        }
        
        public void OCR(bool isOcrOnRectangleArea)
        {
            byte[] sbytes = null;
            string languageFolder = m_strCurrentDirectory + @"OCR\tessdata"; 
            m_Tesseract.TessDataPath = languageFolder;
            m_Tesseract.Language = languages["English"];
            m_Tesseract.ResultFormat = ResultFormat.PDFPlainText;
            string imageFolder = m_strCurrentDirectory + @"OCR\OCRImages";
            //string strfilename = imageFolder +  @"\DNTImage2.tif";
            string path = m_strCurrentDirectory + @"Upload";
            //int pos = strfilename.LastIndexOf(".");
            //if (pos != -1)
            //{
            //    string strSuffix = strfilename.Substring(pos, strfilename.Length - pos).ToLower();
            //    if (strSuffix.CompareTo(".pdf") == 0)
            //    {
            //        m_PDFRasterizer.ConvertMode = Dynamsoft.PDF.Enums.EnumConvertMode.enumCM_AUTO;
            //        m_PDFRasterizer.ConvertToImage(strfilename, "", 200, this as IConvertCallback);                    
            //    }
            //}
            //m_ImageCore.IO.LoadImage(strfilename);
            m_ImageCore.IO.LoadImage(imageFolder + @"\DNTImage1.tif");
            m_ImageCore.IO.LoadImage(imageFolder + @"\DNTImage2.tif");
            m_ImageCore.IO.LoadImage(imageFolder + @"\DNTImage3.tif");
            m_ImageCore.IO.LoadImage(imageFolder + @"\DNTImage4.tif");
            m_ImageCore.IO.LoadImage(imageFolder + @"\DNTImage5.tif");
            m_ImageCore.IO.LoadImage(imageFolder + @"\DNTImage6.tif");
            m_ImageCore.IO.LoadImage(imageFolder + @"\DNTImage7.tif");
            int imgCount = m_ImageCore.ImageBuffer.HowManyImagesInBuffer;
            for (short index = 0; index < imgCount; index++)
            {
                if (index >= 0 && index < m_ImageCore.ImageBuffer.HowManyImagesInBuffer)
                {
                    if (tempListSelectedBitmap == null)
                    {
                        tempListSelectedBitmap = new List<Bitmap>();
                    }
                    Bitmap temp = m_ImageCore.ImageBuffer.GetBitmap(index);
                    tempListSelectedBitmap.Add(temp);
                }
            }
            //Bitmap temp = m_ImageCore.ImageBuffer.GetBitmap(0);
            //tempListSelectedBitmap.Add(temp);
            if (tempListSelectedBitmap != null)
                sbytes = m_Tesseract.Recognize(tempListSelectedBitmap);
            if (sbytes != null && sbytes.Length > 0)
            {
                File.WriteAllBytes(path + @"\" + "Test.pdf", sbytes);                
            }        
        }

    }

    
}