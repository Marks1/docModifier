using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using document_operation;


namespace runner
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args[0] == "help"){
                Console.WriteLine("runner.exe [pdf|doc|excel|ppt] options \n");
                Console.WriteLine("runner.exe pdf htmlFile pdfFile \n");
                Console.WriteLine("runner.exe docx wordFile str url");
            }
            //string pdffile_path = args[0];
            //pdfChecker pdf = new pdfChecker();
            //pdf.ReadPdfFile(pdffile_path);
            //if (pdf.SameAsExpected("a.txt"))
            //{
            //    Console.WriteLine("Same");
            //}
            //else {
            //    Console.WriteLine("not same");
            //}

            //runner.exe htmlTemp pdfFIle

            if (args[0] == "pdf")
            {
                string html = args[1];
                string pdfPath = args[2];
                PdfOperation pdf = new PdfOperation();
            
                pdf.CreatePdf(html, pdfPath);
            }

            if (args[0] == "docx")
            {
                string wordFile = args[1];
                string url = args[3];
                string targetStr = args[2];

                OfficeOperation officeDoc = new OfficeOperation();
                officeDoc.Word2007SearchAndReplace(wordFile, targetStr, url);
            }

            if (args[0] == "xlsx")
            {
                string excelFile = args[1];
                string url = args[3];
                string targetStr = args[2];

                OfficeOperation officeDoc = new OfficeOperation();
                officeDoc.Excel2007SearchAndReplace(excelFile, targetStr, url);
            }

            if (args[0] == "pptx")
            {
                string pptFile = args[1];
                string url = args[3];
                string targetStr = args[2];
                OfficeOperation pptDoc = new OfficeOperation();
                pptDoc.PPT2007SearchAndReplace(pptFile, targetStr, url);
            }
        }
    }
}
