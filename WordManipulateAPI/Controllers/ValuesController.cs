using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Web;
using System.Web.Http;
using WordManipulateAPI.Models;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http.Cors;
using System.Configuration;
using System.Text;
using OpenHtmlToPdf;
using System.Collections.Generic;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Path = System.IO.Path;
using System.Linq;

namespace WordManipulateAPI.Controllers
{
    //[EnableCors(origins: "http://localhost:4200", headers: "*", methods: "*")]
    //[EnableCors(origins: "http://localhost:4200", headers: "*", methods: "*")]
    [EnableCors(origins: "*", headers: "*", methods: "*", SupportsCredentials =true)]
    [Authorize]
    public class ValuesController : ApiController
    {
        static iTextSharp.text.pdf.PdfStamper stamper = null;
        public class TemplateData
        {
            public string HTMLContent { get; set; }
            public string FullFileName { get; set; }
        }
        public class UpdatePDF
        {
            public string ExistingVariableinPDF { get; set; }
            public string ReplacingVariable { get; set; }
            public string ReplacingDateVariable { get; set; }
            public string sourceFile { get; set; }
          
        }
        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/GeneratePDFFileFromHtml")]
        public HttpResponseMessage GeneratePDFFileFromHtml(TemplateData templateData)
        {
            HttpResponseMessage response = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
            try
            {
                var pdf = Pdf.From(templateData.HTMLContent).Content();
                System.IO.File.WriteAllBytes(@templateData.FullFileName, pdf);


            }
            catch (Exception ex)
            {

                response.StatusCode = System.Net.HttpStatusCode.PreconditionFailed;

            }
            finally
            {

            }
            return response;
        }


        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/UpdatePdfDoc")]
        public HttpResponseMessage UpdatePdfDoc(UpdatePDF updatePDF)
        {
            HttpResponseMessage response = null;
            try
            {
                string ExistingVariableinPDF = updatePDF.ExistingVariableinPDF;
                string ReplacingVariable = updatePDF.ReplacingVariable;
                string ReplacingDateVariable = updatePDF.ReplacingDateVariable;
                string sourceFile = @updatePDF.sourceFile;

                FileInfo fileInfo = new FileInfo(sourceFile);
                string tempfilename = Path.Combine(fileInfo.Directory.FullName, ReplacingVariable + fileInfo.Extension);
                string destFile = @tempfilename;

                PdfReader pReader = new PdfReader(sourceFile);
                int numberOfPages = pReader.NumberOfPages;
                FileStream fileStream = new System.IO.FileStream(destFile, System.IO.FileMode.Create);
                stamper = new iTextSharp.text.pdf.PdfStamper(pReader, fileStream);
                PDFTextGetter(ExistingVariableinPDF, ReplacingVariable, StringComparison.CurrentCultureIgnoreCase, sourceFile, destFile);
                if(!String.IsNullOrEmpty(ReplacingDateVariable))
                PDFTextGetter(ReplacingDateVariable, DateTime.Now.DisplayWithSuffix(), StringComparison.CurrentCultureIgnoreCase, sourceFile, destFile);

                stamper.Close();
                stamper.Dispose();
                pReader.Close();
                pReader.Dispose();
                fileStream.Close();
                fileInfo.Refresh();

                File.Delete(sourceFile);
                response = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
            }
            catch (Exception ex)
            {
                response = new HttpResponseMessage(System.Net.HttpStatusCode.InternalServerError);
                Logger.WriteLog("Exception while UpdatePdfDoc " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
            
           // File.Move(tempfilename, sourceFile);
            return response;
        }

        public static void PDFTextGetter(string pSearch, string replacingText, StringComparison SC, string SourceFile, string DestinationFile)
        {
            try
            {
                iTextSharp.text.pdf.PdfContentByte cb = null;
                iTextSharp.text.pdf.PdfContentByte cb2 = null;
                iTextSharp.text.pdf.PdfWriter writer = null;
                iTextSharp.text.pdf.BaseFont bf = null;
                FileInfo path = new FileInfo(ConfigurationManager.AppSettings["FontPath"]);
                iTextSharp.text.FontFactory.RegisterDirectory(path.Directory.FullName);
                if (System.IO.File.Exists(SourceFile))
                {
                    PdfReader pReader = new PdfReader(SourceFile);
                    for (int page = 1; page <= pReader.NumberOfPages; page++)
                    {
                        myLocationTextExtractionStrategy strategy = new myLocationTextExtractionStrategy();
                        string currentText = "";
                        cb = stamper.GetOverContent(page);
                        cb2 = stamper.GetOverContent(page);
                        //Send some data contained in PdfContentByte, looks like the first is always cero for me and the second 100,
                        //but i'm not sure if this could change in some cases
                        strategy.UndercontentCharacterSpacing = (int)cb.CharacterSpacing;
                        strategy.UndercontentHorizontalScaling = (int)cb.HorizontalScaling;
                        //It's not really needed to get the text back, but we have to call this line ALWAYS,
                        //because it triggers the process that will get all chunks from PDF into our strategy Object
                        currentText = PdfTextExtractor.GetTextFromPage(pReader, page, strategy);
                        //The real getter process starts in the following line
                        List<iTextSharp.text.Rectangle> MatchesFound = strategy.GetTextLocations(pSearch, SC);
                        //Set the fill color of the shapes, I don't use a border because it would make the rect bigger
                        //but maybe using a thin border could be a solution if you see the currect rect is not big enough to cover all the text it should cover
                        cb.SetColorFill(iTextSharp.text.BaseColor.WHITE);
                        //MatchesFound contains all text with locations, so do whatever you want with it, this highlights them using PINK color:
                        foreach (iTextSharp.text.Rectangle rect in MatchesFound)
                        {
                            //width
                            cb.Rectangle(rect.Left, rect.Bottom, rect.Width, rect.Height);
                            cb.Fill();
                            cb2.SetColorFill(iTextSharp.text.BaseColor.BLACK);
                            bf = GetFont();
                            //bf = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252 , BaseFont.NOT_EMBEDDED);
                            cb2.SetFontAndSize(bf, 9);
                            cb2.BeginText();
                            cb2.ShowTextAligned(0, replacingText, rect.Left, rect.Bottom + 2, 0);
                            cb2.EndText();
                            cb2.Fill();
                            //break;
                        }
                      //break;
                    }
                    pReader.Close();
                }
            }
            catch (Exception ex)
            {
            }
        }

public static BaseFont GetFont()
{
    return BaseFont.CreateFont(ConfigurationManager.AppSettings["FontPath"], BaseFont.CP1252, BaseFont.EMBEDDED);
}

private void ManipulatePdf(String src, String dest)
        {

            //PdfDocument pdfDoc = new PdfDocument(new PdfReader(src), new PdfWriter(dest));
            //PdfPage page = pdfDoc.GetFirstPage();
            //PdfDictionary dict = page.GetPdfObject();           
            //PdfObject pdfObject = dict.Get(PdfName.Contents);
            //if (pdfObject is PdfStream)
            //{
            //    PdfStream stream = (PdfStream)pdfObject;
            //    byte[] data = stream.GetBytes();
            //    String replacedData = iText.IO.Util.JavaUtil.GetStringForBytes(data).Replace("pune", "PUNE");
            //    stream.SetData((Encoding.UTF8.GetBytes(replacedData)));
            //}

            //pdfDoc.Close();

            string fileNameExisting = @"d:\temp\testpdf\TECH10031673.pdf";
            string fileNameNew = @"d:\temp\testpdf\TECH10031673-new.pdf";

            //variables
            string pathin = fileNameExisting;
            string pathout = fileNameNew;
            PdfReader reader = new PdfReader(pathin);
            PdfStamper stamper = new PdfStamper(reader, new FileStream(pathout, FileMode.Create));
            //select two pages from the original document
            reader.SelectPages("1-1");
            //gettins the page size in order to substract from the iTextSharp coordinates
            var pageSize = reader.GetPageSize(1);
            // PdfContentByte from stamper to add content to the pages over the original content
            PdfContentByte pbover = stamper.GetOverContent(1);
            //pbover.SaveState();
            // for better coorniate, it would be better use pageSize object and calculate from top like text position in the below
            iTextSharp.text.Rectangle rectangle = new iTextSharp.text.Rectangle(1200, 1850, 1530, 1890);
            rectangle.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
            pbover.Rectangle(rectangle);

            BaseFont bf = BaseFont.CreateFont(ConfigurationManager.AppSettings["FontPath"], BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            //Font font = new Font();
            // font.Size = 45;

            //setting up the X and Y coordinates of the document
            int x = 1300;
            int y = 160;

            x += 113;
            y = (int)(pageSize.Height - y);


        }


        public class myLocationTextExtractionStrategy : LocationTextExtractionStrategy
        {

            public float UndercontentCharacterSpacing { get; set; }
            public float UndercontentHorizontalScaling { get; set; }
            private SortedList<string, DocumentFont> ThisPdfDocFonts = new SortedList<string, DocumentFont>();
            private List<TextChunk> locationalResult = new List<TextChunk>();

            private bool StartsWithSpace(String inText)
            {
                if (string.IsNullOrEmpty(inText))
                {
                    return false;
                }
                if (inText.StartsWith(" "))
                {
                    return true;
                }
                return false;
            }

            private bool EndsWithSpace(String inText)
            {
                if (string.IsNullOrEmpty(inText))
                {
                    return false;
                }
                if (inText.EndsWith(" "))
                {
                    return true;
                }
                return false;
            }

            public override string GetResultantText()
            {
                locationalResult.Sort();

                StringBuilder sb = new StringBuilder();
                TextChunk lastChunk = null;

                foreach (var chunk in locationalResult)
                {
                    if (lastChunk == null)
                    {
                        sb.Append(chunk.text);
                    }
                    else
                    {
                        if (chunk.SameLine(lastChunk))
                        {
                            float dist = chunk.DistanceFromEndOf(lastChunk);
                            if (dist < -chunk.charSpaceWidth)
                            {
                                sb.Append(" ");
                            }
                            else if (dist > chunk.charSpaceWidth / 2.0F && !StartsWithSpace(chunk.text) && !EndsWithSpace(lastChunk.text))
                            {
                                sb.Append(" ");
                            }
                            sb.Append(chunk.text);
                        }
                        else
                        {
                            sb.Append("\n");
                            sb.Append(chunk.text);
                        }
                    }
                    lastChunk = chunk;
                }
                return sb.ToString();
            }

            static string UnicodeToUTF8(string from)
            {
                var bytes = Encoding.UTF8.GetBytes(from);
                return new string(bytes.Select(b => (char)b).ToArray());
            }

            public List<iTextSharp.text.Rectangle> GetTextLocations(string pSearchString, System.StringComparison pStrComp)
            {
                List<iTextSharp.text.Rectangle> FoundMatches = new List<iTextSharp.text.Rectangle>();
                StringBuilder sb = new StringBuilder();
                List<TextChunk> ThisLineChunks = new List<TextChunk>();
                bool bStart = false;
                bool bEnd = false;
                TextChunk FirstChunk = null;
                TextChunk LastChunk = null;
                string sTextInUsedChunks = null;

                foreach (var chunk in locationalResult)
                {
                    if (ThisLineChunks.Count > 0 && !chunk.SameLine(ThisLineChunks.Last()))
                    {
                        var input = sb.ToString().Replace((char)173, '-');
                        sb.Clear();
                        sb.Append(input);
                        if (sb.ToString().IndexOf(pSearchString, pStrComp) > -1)
                        {
                            string sLine = sb.ToString();

                            int iCount = 0;
                            int lPos = 0;
                            lPos = sLine.IndexOf(pSearchString, 0, pStrComp);
                            while (lPos > -1)
                            {
                                iCount++;
                                if (lPos + pSearchString.Length > sLine.Length)
                                {
                                    break;
                                }
                                else
                                {
                                    lPos = lPos + pSearchString.Length;
                                }
                                lPos = sLine.IndexOf(pSearchString, lPos, pStrComp);
                            }

                            int curPos = 0;
                            for (int i = 1; i <= iCount; i++)
                            {
                                string sCurrentText;
                                int iFromChar;
                                int iToChar;

                                iFromChar = sLine.IndexOf(pSearchString, curPos, pStrComp);
                                curPos = iFromChar;
                                iToChar = iFromChar + pSearchString.Length - 1;
                                sCurrentText = null;
                                sTextInUsedChunks = null;
                                FirstChunk = null;
                                LastChunk = null;

                                foreach (var chk in ThisLineChunks)
                                {
                                    sCurrentText = sCurrentText + chk.text;

                                    if (!bStart && sCurrentText.Length - 1 >= iFromChar)
                                    {
                                        FirstChunk = chk;
                                        bStart = true;
                                    }

                                    if (bStart && !bEnd)
                                    {
                                        sTextInUsedChunks = sTextInUsedChunks + chk.text;
                                    }

                                    if (!bEnd && sCurrentText.Length - 1 >= iToChar)
                                    {
                                        LastChunk = chk;
                                        bEnd = true;
                                    }
                                    if (bStart && bEnd)
                                    {
                                        FoundMatches.Add(GetRectangleFromText(FirstChunk, LastChunk, pSearchString, sTextInUsedChunks, iFromChar, iToChar, pStrComp));
                                        curPos = curPos + pSearchString.Length;
                                        bStart = false;
                                        bEnd = false;
                                        break;
                                    }
                                }
                            }
                        }
                        sb.Clear();
                        ThisLineChunks.Clear();
                        //if (FoundMatches.Count() > 0)
                        //    break;
                    }
                    if (!String.IsNullOrEmpty(chunk.text.Trim()))
                    {
                        ThisLineChunks.Add(chunk);
                        sb.Append(chunk.text);
                    }
                }
                return FoundMatches;
            }

            private iTextSharp.text.Rectangle GetRectangleFromText(TextChunk FirstChunk, TextChunk LastChunk, string pSearchString,
              string sTextinChunks, int iFromChar, int iToChar, System.StringComparison pStrComp)
            {
                sTextinChunks = sTextinChunks.Replace((char)173, '-');

                float LineRealWidth = LastChunk.PosRight - FirstChunk.PosLeft;

                float LineTextWidth = GetStringWidth(sTextinChunks, LastChunk.curFontSize,
                                                             LastChunk.charSpaceWidth,
                                                             ThisPdfDocFonts.ElementAt(LastChunk.FontIndex).Value);

                float TransformationValue = LineRealWidth / LineTextWidth;

                int iStart = sTextinChunks.IndexOf(pSearchString, pStrComp);

                int iEnd = iStart + pSearchString.Length - 1;

                string sLeft;
                if (iStart == 0)
                {
                    sLeft = null;
                }
                else
                {
                    sLeft = sTextinChunks.Substring(0, iStart);
                }

                string sRight;
                if (iEnd == sTextinChunks.Length - 1)
                {
                    sRight = null;
                }
                else
                {
                    sRight = sTextinChunks.Substring(iEnd + 1, sTextinChunks.Length - iEnd - 1);
                }

                float LeftWidth = 0;
                if (iStart > 0)
                {
                    LeftWidth = GetStringWidth(sLeft, LastChunk.curFontSize,
                                                      LastChunk.charSpaceWidth,
                                                      ThisPdfDocFonts.Values.ElementAt(LastChunk.FontIndex));
                    LeftWidth = LeftWidth * TransformationValue;
                }

                float RightWidth = 0;
                if (iEnd < sTextinChunks.Length - 1)
                {
                    RightWidth = GetStringWidth(sRight, LastChunk.curFontSize,
                                                        LastChunk.charSpaceWidth,
                                                        ThisPdfDocFonts.Values.ElementAt(LastChunk.FontIndex));
                    RightWidth = RightWidth * TransformationValue;
                }

                float LeftOffset = FirstChunk.distParallelStart + LeftWidth;
                float RightOffset = LastChunk.distParallelEnd - RightWidth;
                return new iTextSharp.text.Rectangle(LeftOffset, FirstChunk.PosBottom, RightOffset, FirstChunk.PosTop);
            }

            private float GetStringWidth(string str, float curFontSize, float pSingleSpaceWidth, DocumentFont pFont)
            {

                char[] chars = str.ToCharArray();
                float totalWidth = 0;
                float w = 0;

                foreach (Char c in chars)
                {
                    w = pFont.GetWidth(c) / 1000;
                    totalWidth += (w * curFontSize + this.UndercontentCharacterSpacing) * this.UndercontentHorizontalScaling / 100;
                }

                return totalWidth;
            }

            public override void RenderText(TextRenderInfo renderInfo)
            {
                LineSegment segment = renderInfo.GetBaseline();
                TextChunk location = new TextChunk(renderInfo.GetText(), segment.GetStartPoint(), segment.GetEndPoint(), renderInfo.GetSingleSpaceWidth());

                location.PosLeft = renderInfo.GetDescentLine().GetStartPoint()[Vector.I1];
                location.PosRight = renderInfo.GetAscentLine().GetEndPoint()[Vector.I1];
                location.PosBottom = renderInfo.GetDescentLine().GetStartPoint()[Vector.I2];
                location.PosTop = renderInfo.GetAscentLine().GetEndPoint()[Vector.I2];
                location.curFontSize = location.PosTop - segment.GetStartPoint()[Vector.I2];

                string StrKey = renderInfo.GetFont().PostscriptFontName + location.curFontSize.ToString();
                if (!ThisPdfDocFonts.ContainsKey(StrKey))
                {
                    ThisPdfDocFonts.Add(StrKey, renderInfo.GetFont());
                }
                location.FontIndex = ThisPdfDocFonts.IndexOfKey(StrKey);
                locationalResult.Add(location);
            }

            private class TextChunk : IComparable<TextChunk>
            {
                public string text { get; set; }
                public Vector startLocation { get; set; }
                public Vector endLocation { get; set; }
                public Vector orientationVector { get; set; }
                public int orientationMagnitude { get; set; }
                public int distPerpendicular { get; set; }
                public float distParallelStart { get; set; }
                public float distParallelEnd { get; set; }
                public float charSpaceWidth { get; set; }
                public float PosLeft { get; set; }
                public float PosRight { get; set; }
                public float PosTop { get; set; }
                public float PosBottom { get; set; }
                public float curFontSize { get; set; }
                public int FontIndex { get; set; }


                public TextChunk(string str, Vector startLocation, Vector endLocation, float charSpaceWidth)
                {
                    this.text = str;
                    this.startLocation = startLocation;
                    this.endLocation = endLocation;
                    this.charSpaceWidth = charSpaceWidth;

                    Vector oVector = endLocation.Subtract(startLocation);
                    if (oVector.Length == 0)
                    {
                        oVector = new Vector(1, 0, 0);
                    }
                    orientationVector = oVector.Normalize();
                    orientationMagnitude = (int)(Math.Truncate(Math.Atan2(orientationVector[Vector.I2], orientationVector[Vector.I1]) * 1000));

                    Vector origin = new Vector(0, 0, 1);
                    distPerpendicular = (int)((startLocation.Subtract(origin)).Cross(orientationVector)[Vector.I3]);

                    distParallelStart = orientationVector.Dot(startLocation);
                    distParallelEnd = orientationVector.Dot(endLocation);
                }

                public bool SameLine(TextChunk a)
                {
                    var Value = true;
                    if (orientationMagnitude != a.orientationMagnitude)
                    {
                        Value = false;
                        return Value;
                    }
                    if (distPerpendicular != a.distPerpendicular)
                    {
                        Value = false;
                        return Value;
                    }
                    return Value;
                }

                public float DistanceFromEndOf(TextChunk other)
                {
                    float distance = distParallelStart - other.distParallelEnd;
                    return distance;
                }

                int IComparable<TextChunk>.CompareTo(TextChunk rhs)
                {
                    if (this == rhs)
                    {
                        return 0;
                    }

                    int rslt;
                    rslt = orientationMagnitude.CompareTo(rhs.orientationMagnitude);
                    if (rslt != 0)
                    {
                        return rslt;
                    }

                    rslt = distPerpendicular.CompareTo(rhs.distPerpendicular);
                    if (rslt != 0)
                    {
                        return rslt;
                    }
                    rslt = (distParallelStart < rhs.distParallelStart ? -1 : 1);

                    return rslt;
                }
            }

        }


        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("api/Check")]
        public HttpResponseMessage Check()
        {
            HttpResponseMessage response = null;
            try
            {
                Logger.WriteLog("All good");
                response = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
            }
            catch (Exception ex)
            {
                response = new HttpResponseMessage(System.Net.HttpStatusCode.InternalServerError);
                Logger.WriteLog("Exception while UpdatePdfDoc " + ex.Message + Environment.NewLine + ex.StackTrace);
            }

            // File.Move(tempfilename, sourceFile);
            return response;
        }

    }

    public static class extention
    {
        public static string DisplayWithSuffix(this DateTime givenDate)
        {
            int myday;
            string strsuff = string.Empty;
            string mynewdate = string.Empty;
            strsuff = "th";
            myday = givenDate.Day;
            if (myday == 1 | myday == 21 | myday == 31)
                strsuff = "st";
            if (myday == 2 | myday == 22)
                strsuff = "nd";
            if (myday == 3 | myday == 23)
                strsuff = "rd";
            mynewdate = Convert.ToString(myday + "<sup>" + strsuff + "</sup> " + givenDate.ToString("MMMM yyyy"));
            return mynewdate;
        }
    }
}
