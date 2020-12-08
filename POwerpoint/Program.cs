using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Microsoft.VisualBasic;
using System.IO;

using Draftable.CompareAPI.Client;

namespace POwerpoint
{
    class Program
    {
        static Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

        static void Main(string[] args)
        {

            string memofile = converttopdf("C:\\ITTesting\\PowerPoint\\Concepts\\Memmo.pptx");
            string templatefile = converttopdf("C:\\ITTesting\\PowerPoint\\Concepts\\Response.pptx");
            Document memodocument = GetDocument(memofile);
            Document templatedocument = GetDocument(templatefile);
            Document comparedcument = compared(memodocument, templatedocument);
            int score = CalculateScore(comparedcument);
            Console.WriteLine("The Marking Is Finished here  is your  score {0}", score );
            Console.ReadKey();
        
        }
      
        public static Document GetDocument(string path)
        {
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            wordApp.Options.DoNotPromptForConvert = true;
            wordApp.Options.ConfirmConversions = false;
            return wordApp.Documents.OpenNoRepairDialog(path, false, true);
        }
     
        public static Document compared( Document memo, Document template)
        {
          
            var ComparedDocuments = wordApp.CompareDocuments(template, memo);
            return ComparedDocuments;
        }
        public static string GenerateFileName()
        {
            var date = DateTime.Now;
            return "C:\\ITTesting\\PowerPoint\\converted" + date.Hour + date.Minute + date.Second + date.Millisecond + ".pdf";
        }
        public static string converttopdf(string filename)
        {
            var app = new Microsoft.Office.Interop.PowerPoint.Application();
            app.Visible = MsoTriState.msoFalse;
            var pres = app.Presentations;
            string pdfFileName = GenerateFileName();
            var file = pres.Open(filename, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoTrue);
            file.SaveAs(pdfFileName, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF, Microsoft.Office.Core.MsoTriState.msoTrue);
            return pdfFileName;
        }
        public static int CalculateScore(Document ComparedDocuments)
        {
            int TotalPoints = 100;
            int PenaltyPoint = 2;
            int revisioncounter = 0;
            int deletionOrinsetioncounter = 0;
            int moveToOrFromCounter = 0;
            foreach (Section s in ComparedDocuments.Sections)
            {
                foreach (Revision revision in s.Range.Revisions)
                {
                    if ((revision.Type == WdRevisionType.wdRevisionProperty) || (revision.Type == WdRevisionType.wdRevisionParagraphProperty) || (revision.Type == WdRevisionType.wdRevisionSectionProperty))
                    {
                        revisioncounter++;
                    }
                    if ((revision.Type == WdRevisionType.wdRevisionDelete) || (revision.Type == WdRevisionType.wdRevisionInsert))
                    {
                        deletionOrinsetioncounter++;
                    }
                    if ((revision.Type == WdRevisionType.wdRevisionMovedFrom) || (revision.Type == WdRevisionType.wdRevisionMovedTo))
                    {
                        moveToOrFromCounter++;
                    }
                    string revisioncontent = revision.Range.Text;
                }
            }
            TotalPoints -= ((revisioncounter + moveToOrFromCounter + deletionOrinsetioncounter) * PenaltyPoint);
            if (TotalPoints < 0)
            {
                TotalPoints = 0;
            }
            return TotalPoints;
        }
    }
    
}
