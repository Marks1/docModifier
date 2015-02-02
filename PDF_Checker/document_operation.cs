using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
//using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Presentation;
using OD = DocumentFormat.OpenXml.Drawing;

namespace document_operation
{
    public class PdfOperation
    {
        protected string extractedContent = "";

        public string ReadPdfFile(string fileName)
        {
            StringBuilder text = new StringBuilder();

            if (File.Exists(fileName))
            {
                PdfReader pdfReader = new PdfReader(fileName);

                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                }
                pdfReader.Close();
            }
            return text.ToString();
        }

        public void CreatePdf(string HtmlName, string PdfName)
        {
            //PdfName = "sample.pdf";

            iTextSharp.text.Rectangle rect = PageSize.LETTER;
            iTextSharp.text.Document doc = new iTextSharp.text.Document(PageSize.A4);
            var output = new FileStream(PdfName, FileMode.Create);
            var writer = PdfWriter.GetInstance(doc, output);

            doc.Open();


            string htmlContent = System.IO.File.ReadAllText(HtmlName);
            List<IElement> ie = HTMLWorker.ParseToList(new StringReader(htmlContent), null);
   
            float pageWidth = rect.Width;

            foreach (IElement element in ie)
            {
                PdfPTable table = element as PdfPTable;

                if (table != null)
                {
                    table.SetWidthPercentage(
                        new float[] {
                        (float).25 * pageWidth, 
                        (float).50 * pageWidth, 
                        (float).25 * pageWidth},
                        rect
                    );
                }
                doc.Add(element);
            }
            doc.Close();  
           
        }
    

    }

    public class OfficeOperation
    {

        public void Word2007SearchAndReplace(string srcWordName, string searchStr, string replaceStr)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(srcWordName, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex(searchStr);
                docText = regexText.Replace(docText, replaceStr);

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }

        }

        public void Excel2007SearchAndReplace(string srcExcelName, string searchStr, string replaceStr)
        {
            FileInfo newFile = new FileInfo(srcExcelName);
            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
            {
                string content = null;
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[1];

                content = worksheet.Cell(1, 1).Value;
                Regex regexText = new Regex(searchStr);
                content = regexText.Replace(content, replaceStr);
                worksheet.Cell(1, 1).Value = content;
                try
                {
                    xlPackage.Save();
                }
                catch
                {
                    return;
                }
            }
            
        }

        public void PPT2007SearchAndReplace(string srcPPTName, string searchStr, string replaceStr)
        {
            using (PresentationDocument prstDoc = PresentationDocument.Open(srcPPTName, true))
            {

                Slide firstSlide = prstDoc.PresentationPart.SlideParts.ElementAt(0).Slide;

                Shape firstShape = firstSlide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().ElementAt(0);

                OD.Paragraph para = firstShape.TextBody.ChildElements.OfType<OD.Paragraph>().ElementAt(0);

                int cnt = para.ChildElements.OfType<OD.Run>().Count();

                OD.Text t;
                Regex regexText = new Regex(searchStr);
                for (int i = 0; i < cnt; i++)
                {
                    t = para.ChildElements.OfType<OD.Run>().ElementAt(i).Text;
                    t.Text = regexText.Replace(t.Text, replaceStr);
                }

                //t.Text = DateTime.Now.ToString("mm/dd/yyyy");   // if you want to change the format you can change here.

                prstDoc.PresentationPart.Presentation.Save();

            }
        }
    }
}


