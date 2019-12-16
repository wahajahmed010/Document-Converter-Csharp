using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace Document_converter
{
    class PdfToWord : Conversion
    {
        public void Converter(string fopen, string fsave)
        {
            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
            f.OpenPdf(fopen);
            //f.ToWord(textBox2.Text);

            if (f.PageCount > 0)
            {
                // You may choose output format between Docx and Rtf.
                f.WordOptions.Format = SautinSoft.PdfFocus.CWordOptions.eWordDocument.Docx;

                int result = f.ToWord(fsave);
                

                // Show the resulting Word document.
                if (result == 0)
                {
                    //Form1 main = new Form1();
                   
                  //  main.popup();

                    //System.Diagnostics.Process.Start(textBox2.Text);
                    MessageBox.Show("Conversion is Completed!");
                }

                

            }
        }
    }
}
