using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace WriteWord
{
    class WordToPDF
    {
        _Application wordApp;
        Object Nothing;

        public WordToPDF()
        {
            wordApp = new ApplicationClass();
            Nothing = Missing.Value;
        }

        public void ChangeToPDF(object WordPath, string PDFPath)
        {
            Console.WriteLine("开始将Word转成PDF");
            _Document doc = wordApp.Documents.Open(ref WordPath, ref Nothing);
            doc.ExportAsFixedFormat(PDFPath, WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, true, true, WdExportCreateBookmarks.wdExportCreateNoBookmarks, true, true, false);
            object NotSave = WdSaveOptions.wdDoNotSaveChanges;
            doc.Close(ref NotSave, ref Nothing, ref Nothing);
        }

        public void Close()
        {
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }
    }
}
