using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Document_converter
{
    class PdfToTxt : Conversion, IDisposable
    {
        private SautinSoft.PdfFocus _pdfFocus;
        private bool _disposed;

        public void Convert(string fopen, string fsave, IProgress<double> progress, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(fopen))
                throw new ArgumentException("Input path cannot be empty", nameof(fopen));
            if (!File.Exists(fopen))
                throw new FileNotFoundException("Input file not found", fopen);
            if (string.IsNullOrWhiteSpace(fsave))
                throw new ArgumentException("Output path cannot be empty", nameof(fsave));

            string pdfFile = fopen;
            string outputPath = fsave;

            token.ThrowIfCancellationRequested();
            progress?.Report(0.1);

            _pdfFocus = new SautinSoft.PdfFocus();
            _pdfFocus.OpenPdf(pdfFile);

            token.ThrowIfCancellationRequested();
            progress?.Report(0.5);

            if (_pdfFocus.PageCount > 0)
            {
                token.ThrowIfCancellationRequested();
                int result = _pdfFocus.ToText(outputPath);

                token.ThrowIfCancellationRequested();
                progress?.Report(0.9);

                if (result == 0)
                {
                    MessageBox.Show("Conversion is Completed!");
                }
            }
        }

        public void Converter(string filepath, string savepath)
        {
            if (string.IsNullOrWhiteSpace(filepath))
                throw new ArgumentException("Input path cannot be empty", nameof(filepath));
            if (!File.Exists(filepath))
                throw new FileNotFoundException("Input file not found", filepath);
            if (string.IsNullOrWhiteSpace(savepath))
                throw new ArgumentException("Output path cannot be empty", nameof(savepath));

            string pdfFile = filepath;
            string outputPath = savepath;

            _pdfFocus = new SautinSoft.PdfFocus();
            _pdfFocus.OpenPdf(pdfFile);

            if (_pdfFocus.PageCount > 0)
            {
                int result = _pdfFocus.ToText(outputPath);

                if (result == 0)
                {
                    MessageBox.Show("Conversion is Completed!");
                }
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_pdfFocus != null)
                {
                    _pdfFocus.Dispose();
                }
                _disposed = true;
            }
        }
    }
}
