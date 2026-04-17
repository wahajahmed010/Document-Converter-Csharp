using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Document_converter
{
    public partial class Form1 : Form
    {
        private CancellationTokenSource _cts;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Logger.Instance.Info("Application started");
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

        private async void ConvertButton_Click(object sender, EventArgs e)
        {
            openFileDialog3.CheckFileExists = true;
            openFileDialog3.CheckPathExists = true;

            string filepath = FileOpen.Text;
            string savepath = FileSave.Text;

            if (string.IsNullOrWhiteSpace(filepath) || string.IsNullOrWhiteSpace(savepath))
            {
                MessageBox.Show("Please select both input and output files.");
                return;
            }

            string openwali = Path.GetExtension(openFileDialog3.FileName);
            string savewali = Path.GetExtension(saveFileDialog3.FileName);

            if (openwali.Equals(savewali, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Error: Same file extension selected");
                return;
            }

            // Setup cancellation
            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            var progress = new Progress<double>(p => progressBar1.Value = (int)(p * 100));

            try
            {
                Logger.Instance.Info($"Starting conversion: {filepath} -> {savepath}");
                statusLabel.Text = "Converting...";
                progressBar1.Value = 0;
                ConvertButton.Enabled = false;

                var conversion = new Conversion();
                await conversion.ConvertAsync(filepath, savepath, openwali, savewali, progress, token);

                Logger.Instance.Info("Conversion completed successfully");
                MessageBox.Show("Conversion Completed!");
                progressBar1.Value = 100;
            }
            catch (OperationCanceledException)
            {
                Logger.Instance.Info("Conversion cancelled by user");
                MessageBox.Show("Conversion cancelled.");
                progressBar1.Value = 0;
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("Conversion failed", ex);
                MessageBox.Show($"Conversion failed: {ex.Message}");
                progressBar1.Value = 0;
            }
            finally
            {
                ConvertButton.Enabled = true;
                statusLabel.Text = "Ready";
                _cts?.Dispose();
                _cts = null;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string file = FileSave.Text;
            if (string.IsNullOrEmpty(file))
            {
                MessageBox.Show("No File Converted Yet!");
            }
            else
            {
                System.Diagnostics.Process.Start(file);
            }
        }

        private void howToUseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 frm2 = new Form2();
            frm2.Show();
        }

        private void aboutAppToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.Show();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void saveFileDialog3_FileOk(object sender, CancelEventArgs e)
        {
        }

        private void openFileDialog3_FileOk(object sender, CancelEventArgs e)
        {
        }
    }
}