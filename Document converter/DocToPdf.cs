using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace Document_converter
{
    class DocToPdf : Conversion, IDisposable
    {
        private Microsoft.Office.Interop.Word.Application _app;
        private Microsoft.Office.Interop.Word.Document _wordDoc;
        private bool _disposed;

        public void Convert(string fopen, string fsave, IProgress<double> progress, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(fopen))
                throw new ArgumentException("Input path cannot be empty", nameof(fopen));
            if (!File.Exists(fopen))
                throw new FileNotFoundException("Input file not found", fopen);
            if (string.IsNullOrWhiteSpace(fsave))
                throw new ArgumentException("Output path cannot be empty", nameof(fsave));

            _app = null;
            _wordDoc = null;

            try
            {
                token.ThrowIfCancellationRequested();
                progress?.Report(0.1);

                _app = new Microsoft.Office.Interop.Word.Application();
                _wordDoc = _app.Documents.Open(fopen);

                token.ThrowIfCancellationRequested();
                progress?.Report(0.5);

                _wordDoc.ExportAsFixedFormat(fsave, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

                token.ThrowIfCancellationRequested();
                progress?.Report(0.9);
            }
            finally
            {
                if (_wordDoc != null)
                {
                    _wordDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.FinalReleaseComObject(_wordDoc);
                }
                if (_app != null)
                {
                    _app.Quit();
                    Marshal.FinalReleaseComObject(_app);
                }
            }
        }

        public void Convertor(string loadpath, string savepath)
        {
            if (string.IsNullOrWhiteSpace(loadpath))
                throw new ArgumentException("Input path cannot be empty", nameof(loadpath));
            if (!File.Exists(loadpath))
                throw new FileNotFoundException("Input file not found", loadpath);
            if (string.IsNullOrWhiteSpace(savepath))
                throw new ArgumentException("Output path cannot be empty", nameof(savepath));

            _app = null;
            _wordDoc = null;

            try
            {
                _app = new Microsoft.Office.Interop.Word.Application();
                _wordDoc = _app.Documents.Open(loadpath);
                _wordDoc.ExportAsFixedFormat(savepath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            }
            finally
            {
                if (_wordDoc != null)
                {
                    _wordDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.FinalReleaseComObject(_wordDoc);
                }
                if (_app != null)
                {
                    _app.Quit();
                    Marshal.FinalReleaseComObject(_app);
                }
            }

            MessageBox.Show("Conversion is Completed!");
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_wordDoc != null)
                {
                    _wordDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.FinalReleaseComObject(_wordDoc);
                }
                if (_app != null)
                {
                    _app.Quit();
                    Marshal.FinalReleaseComObject(_app);
                }
                _disposed = true;
            }
        }
    }
}
