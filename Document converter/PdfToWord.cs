using System;
using System.Threading;
using System.Threading.Tasks;

namespace Document_converter
{
    class PdfToWord
    {
        public void Convert(string fopen, string fsave, IProgress<double> progress, CancellationToken token)
        {
            Logger.Instance.Info($"Converting PDF to Word: {fopen} -> {fsave}");

            token.ThrowIfCancellationRequested();
            progress?.Report(0.1);

            var f = new SautinSoft.PdfFocus();
            f.OpenPdf(fopen);

            token.ThrowIfCancellationRequested();

            if (f.PageCount > 0)
            {
                token.ThrowIfCancellationRequested();
                progress?.Report(0.5);

                f.WordOptions.Format = SautinSoft.PdfFocus.CWordOptions.eWordDocument.Docx;
                int result = f.ToWord(fsave);

                token.ThrowIfCancellationRequested();
                progress?.Report(0.9);

                if (result == 0)
                {
                    Logger.Instance.Info("PDF to Word conversion completed successfully");
                }
                else
                {
                    Logger.Instance.Error($"PDF to Word conversion failed with result: {result}", null);
                }
            }
        }
    }
}