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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        

       // Form2 frm2 = new Form2();
       // Form3 frm3 = new Form3();

        private void button1_Click(object sender, EventArgs e)
        {
            //openFileDialog1.ShowDialog(); 
            //textBox1.Text = openFileDialog1.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //saveFileDialog1.ShowDialog();
           // saveFileDialog1.Filter = "Word Files|*.doc|" + "Docs Files|*.docx|" + "Text files (*.txt)|*.txt";
            //saveFileDialog1.FilterIndex = 1;
            //textBox2.Text = saveFileDialog1.FileName;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            #region Codes Moved

            //string fileopen = textBox1.Text;
            //string filesave = textBox2.Text;
            
            //PdfToWord ptw = new PdfToWord();
            //ptw.Converter(fileopen, filesave);
            
            //SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
            //f.OpenPdf(textBox1.Text);
            ////f.ToWord(textBox2.Text);

            //if (f.PageCount > 0)
            //{
            //    // You may choose output format between Docx and Rtf.
            //    f.WordOptions.Format = SautinSoft.PdfFocus.CWordOptions.eWordDocument.Docx;

            //    int result = f.ToWord(textBox2.Text);

            //    // Show the resulting Word document.
            //    if (result == 0)
            //    {
            //        //System.Diagnostics.Process.Start(textBox2.Text);
            //        MessageBox.Show("Conversion is Completed!");
            //    }
            //}
#endregion 



        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
          
            

        }

        private void button6_Click(object sender, EventArgs e)
        {
            //string filepath = textBox3.Text;
            //string savepath = textBox4.Text;
            //DocToPdf d2p = new DocToPdf();
            //d2p.Convertor(filepath, savepath);
            ////this.timer1.Start();

            
        }

        private void button4_Click(object sender, EventArgs e)
        {

            //openFileDialog2.ShowDialog();
            //textBox3.Text = openFileDialog2.FileName;
            
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //saveFileDialog2.ShowDialog();
            //textBox4.Text = saveFileDialog2.FileName;

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //this.progressBar1.Increment(1);
        }
        public void popup()
        {
            MessageBox.Show("Conversion is Completed!");
        }

        public void Browsebutton_Click(object sender, EventArgs e)
        {
            openFileDialog3.ShowDialog();
            FileOpen.Text = openFileDialog3.FileName;

            
            
        }

        public void SaveButton_Click(object sender, EventArgs e)
        {
            saveFileDialog3.ShowDialog();
            FileSave.Text = saveFileDialog3.FileName;
            
        }

        private void ConvertButton_Click(object sender, EventArgs e)
        {
            //Checking the files that are to be converted.

            openFileDialog3.CheckFileExists = true;
            openFileDialog3.CheckPathExists = true;

            string filepath = FileOpen.Text;
            string savepath = FileSave.Text;
            string filename = Path.GetFileName(filepath);
            string sfilename = Path.GetFileName(savepath);

            string openwali = Path.GetExtension(openFileDialog3.FileName);
            string savewali = Path.GetExtension(saveFileDialog3.FileName);

            Conversion convertchecker = new Conversion();
            convertchecker.fileconvert(filepath, savepath, openwali, savewali, filename, sfilename);

            button1.Visible = true;


            #region file extension codes(moved)
            /////CODES MOVED TO NEW CLASS///

           // //int a = openFileDialog3.FilterIndex;
           // //int b = saveFileDialog3.FilterIndex;
            
           //// MessageBox.Show("Error: Same extension selected");
            
           //     if (openwali == ".doc") //.doc
           //     {
           //         if (savewali == ".docx") //.docx
           //         {
           //             MessageBox.Show("Error: Same file extension");
           //         }
           //         else if (savewali == ".doc") //.doc
           //         {
           //             MessageBox.Show("Error: Same file extension");
           //         }
           //         else if (savewali == ".pdf") //.pdf
           //         {
                        
           //             DocToPdf d2p = new DocToPdf();
           //             d2p.Convertor(filepath, savepath);

           //         }
           //         else if (savewali == ".txt") //.txt
           //         {
           //             //Khaali
           //             MessageBox.Show("Abhi rukja");
           //         }

                    

           //     }
           //     else if (openwali == ".docx") //.docx
           //     {
           //         if (savewali == ".docx") //.docx
           //         {
           //             MessageBox.Show("Error: Same file extension");
           //         }
           //         else if (savewali == ".doc") //.doc
           //         {
           //             MessageBox.Show("Error: Same file extension");
           //         }
           //         else if (savewali == ".pdf") //.pdf
           //         {
                        
           //             DocToPdf d2p = new DocToPdf();
           //             d2p.Convertor(filepath, savepath);

           //         }
           //         else if (savewali == ".txt") //.txt
           //         {
           //             //Khaali
           //             MessageBox.Show("Abhi rukja");
           //         }

           //     }
           //     else if (openwali == ".pdf") //.pdf
           //     {
           //         if (savewali == ".docx") //.docx
           //         {
           //             PdfToWord ptw = new PdfToWord();
           //             ptw.Converter(filepath, savepath);
                        
           //         }
           //         else if (savewali == ".doc") //.doc
           //         {
           //             PdfToWord ptw = new PdfToWord();
           //             ptw.Converter(filepath, savepath);
           //         }
           //         else if (savewali == ".pdf") //.pdf
           //         {
           //             MessageBox.Show("Error: Same file extension");

           //         }
           //         else if (savewali == ".txt") //.txt
           //         {
           //             PdfToTxt PtT = new PdfToTxt();
           //             PtT.Converter(filepath);
           //         }

           //     }
            //     button1.Visible = true;
            #endregion

        }

        private void FileOpen_TextChanged(object sender, EventArgs e)
        {

        }

        private void FileSave_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
            string file = FileSave.Text;
            if (file == null)
            {
                MessageBox.Show("No File Converted Yet!");
            }
            else
            {
                System.Diagnostics.Process.Start(file);
            }

        }

        private void saveFileDialog3_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void openFileDialog3_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void howToUseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //object created on top..
            Form2 frm2 = new Form2();
            frm2.Show();
        }

        private void aboutAppToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //object created on top..
            Form3 frm3 = new Form3();
            frm3.Show();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        
    }
}
