using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Document_converter
{
    class PdfToTxt
    {
        public void Convert(string path, IProgress<double> progress, CancellationToken token)
        {
            Logger.Instance.Info($"Converting PDF to TXT: {path}");

            token.ThrowIfCancellationRequested();
            progress?.Report(0.1);

            string pdfFile = path;
            string textFile = Path.ChangeExtension(pdfFile, ".txt");

            var f = new SautinSoft.PdfFocus();
            f.OpenPdf(pdfFile);

            token.ThrowIfCancellationRequested();
            progress?.Report(0.5);

            if (f.PageCount > 0)
            {
                int result = f.ToText(textFile);

                token.ThrowIfCancellationRequested();
                progress?.Report(0.9);

                if (result == 0)
                {
                    Logger.Instance.Info("PDF to TXT conversion completed successfully");
                }
                else
                {
                    Logger.Instance.Error($"PDF to TXT conversion failed with result: {result}", null);
                }
            }
        }
    }
}