using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Document_converter
{
    class DocToPdf
    {
        public void Convert(string loadpath, string savepath, IProgress<double> progress, CancellationToken token)
        {
            Logger.Instance.Info($"Converting DOC to PDF: {loadpath} -> {savepath}");

            token.ThrowIfCancellationRequested();
            progress?.Report(0.1);

            var app = new Microsoft.Office.Interop.Word.Application();
            var wordDoc = app.Documents.Open(loadpath);

            token.ThrowIfCancellationRequested();
            progress?.Report(0.5);

            wordDoc.ExportAsFixedFormat(savepath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

            progress?.Report(0.9);

            wordDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            app.Quit();

            Marshal.ReleaseComObject(wordDoc);
            Marshal.ReleaseComObject(app);

            Logger.Instance.Info("DOC to PDF conversion completed");
        }
    }
}