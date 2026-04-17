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
using Word = Microsoft.Office.Interop.Word;

namespace Document_converter
{
    class DocToTxt : Conversion, IDisposable
    {
        private Word.Application _app;
        private Word.Document _doc;
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
            _doc = null;

            try
            {
                token.ThrowIfCancellationRequested();
                progress?.Report(0.1);

                object path = fopen;
                object missing = Type.Missing;
                object readOnly = true;
                _app = new Word.Application();
                _doc = _app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                token.ThrowIfCancellationRequested();
                progress?.Report(0.5);

                string text = _doc.Content.Text;
                File.WriteAllText(fsave, text);

                token.ThrowIfCancellationRequested();
                progress?.Report(0.9);
            }
            catch
            {
                MessageBox.Show("An error occured. Please check the file path to your word document, and whether the word document is valid.");
            }
            finally
            {
                if (_doc != null)
                {
                    object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                    _doc.Close(ref saveChanges, ref missing, ref missing);
                    Marshal.FinalReleaseComObject(_doc);
                }
                if (_app != null)
                {
                    _app.Quit(ref missing, ref missing, ref missing);
                    Marshal.FinalReleaseComObject(_app);
                }
            }
        }

        public void Converter(string docPath, string savePath)
        {
            if (string.IsNullOrWhiteSpace(docPath))
                throw new ArgumentException("Input path cannot be empty", nameof(docPath));
            if (!File.Exists(docPath))
                throw new FileNotFoundException("Input file not found", docPath);
            if (string.IsNullOrWhiteSpace(savePath))
                throw new ArgumentException("Output path cannot be empty", nameof(savePath));

            _app = null;
            _doc = null;

            try
            {
                object path = docPath;
                object missing = Type.Missing;
                object readOnly = true;
                _app = new Word.Application();
                _doc = _app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                string text = _doc.Content.Text;
                File.WriteAllText(savePath, text);
            }
            catch
            {
                MessageBox.Show("An error occured. Please check the file path to your word document, and whether the word document is valid.");
            }
            finally
            {
                if (_doc != null)
                {
                    object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                    _doc.Close(ref saveChanges, ref missing, ref missing);
                    Marshal.FinalReleaseComObject(_doc);
                }
                if (_app != null)
                {
                    _app.Quit(ref missing, ref missing, ref missing);
                    Marshal.FinalReleaseComObject(_app);
                }
            }

            MessageBox.Show("Conversion is Completed!");
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_doc != null)
                {
                    object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                    object missing = Type.Missing;
                    _doc.Close(ref saveChanges, ref missing, ref missing);
                    Marshal.FinalReleaseComObject(_doc);
                }
                if (_app != null)
                {
                    object missing = Type.Missing;
                    _app.Quit(ref missing, ref missing, ref missing);
                    Marshal.FinalReleaseComObject(_app);
                }
                _disposed = true;
            }
        }
    }
}
