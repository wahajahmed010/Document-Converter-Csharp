using System;
using System.Threading;
using System.Threading.Tasks;

namespace Document_converter
{
    class Conversion
    {
        public async Task ConvertAsync(string filepath, string savepath, string openwali, string savewali,
            IProgress<double> progress, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            await Task.Run(() =>
            {
                // Handle DOC to PDF/TXT
                if (openwali == ".doc")
                {
                    if (savewali == ".pdf")
                    {
                        var converter = new DocToPdf();
                        converter.Convert(filepath, savepath, progress, token);
                    }
                    else if (savewali == ".txt")
                    {
                        var converter = new DocToTxt();
                        converter.Convert(filepath, savepath, progress, token);
                    }
                }
                // Handle DOCX to PDF/TXT
                else if (openwali == ".docx")
                {
                    if (savewali == ".pdf")
                    {
                        var converter = new DocToPdf();
                        converter.Convert(filepath, savepath, progress, token);
                    }
                    else if (savewali == ".txt")
                    {
                        var converter = new DocToTxt();
                        converter.Convert(filepath, savepath, progress, token);
                    }
                }
                // Handle PDF to DOCX/DOC/TXT
                else if (openwali == ".pdf")
                {
                    if (savewali == ".docx" || savewali == ".doc")
                    {
                        var converter = new PdfToWord();
                        converter.Convert(filepath, savepath, progress, token);
                    }
                    else if (savewali == ".txt")
                    {
                        var converter = new PdfToTxt();
                        converter.Convert(filepath, progress, token);
                    }
                }

                token.ThrowIfCancellationRequested();
                progress?.Report(1.0);
            }, token);
        }
    }
}