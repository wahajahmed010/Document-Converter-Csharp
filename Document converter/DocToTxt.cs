using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Document_converter
{
    class DocToTxt
    {
        public void Convert(string docPath, string savePath, IProgress<double> progress, CancellationToken token)
        {
            Logger.Instance.Info($"Converting DOC to TXT: {docPath} -> {savePath}");

            token.ThrowIfCancellationRequested();
            progress?.Report(0.1);

            object path = docPath;
            Word.Application app = new Word.Application();
            Word.Document doc = null;
            object missing = Type.Missing;
            object readOnly = true;

            try
            {
                token.ThrowIfCancellationRequested();
                doc = app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                token.ThrowIfCancellationRequested();
                progress?.Report(0.5);

                string text = doc.Content.Text;
                File.WriteAllText(savePath, text);

                token.ThrowIfCancellationRequested();
                progress?.Report(0.9);

                Logger.Instance.Info("DOC to TXT conversion completed successfully");
            }
            finally
            {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                app.Quit(ref saveChanges, ref missing, ref missing);
            }
        }
    }
}