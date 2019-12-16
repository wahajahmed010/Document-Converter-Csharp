namespace Document_converter
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.FileOpen = new System.Windows.Forms.TextBox();
            this.FileSave = new System.Windows.Forms.TextBox();
            this.Browsebutton = new System.Windows.Forms.Button();
            this.SaveButton = new System.Windows.Forms.Button();
            this.openFileDialog3 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog3 = new System.Windows.Forms.SaveFileDialog();
            this.ConvertButton = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.homeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.howToUseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutAppToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // FileOpen
            // 
            this.FileOpen.Location = new System.Drawing.Point(29, 54);
            this.FileOpen.Name = "FileOpen";
            this.FileOpen.Size = new System.Drawing.Size(235, 20);
            this.FileOpen.TabIndex = 7;
            this.FileOpen.TextChanged += new System.EventHandler(this.FileOpen_TextChanged);
            // 
            // FileSave
            // 
            this.FileSave.Location = new System.Drawing.Point(29, 105);
            this.FileSave.Name = "FileSave";
            this.FileSave.Size = new System.Drawing.Size(235, 20);
            this.FileSave.TabIndex = 8;
            this.FileSave.TextChanged += new System.EventHandler(this.FileSave_TextChanged);
            // 
            // Browsebutton
            // 
            this.Browsebutton.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.Browsebutton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.Browsebutton.Font = new System.Drawing.Font("Trebuchet MS", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Browsebutton.Location = new System.Drawing.Point(293, 45);
            this.Browsebutton.Name = "Browsebutton";
            this.Browsebutton.Size = new System.Drawing.Size(75, 34);
            this.Browsebutton.TabIndex = 9;
            this.Browsebutton.Text = "Browse";
            this.Browsebutton.UseVisualStyleBackColor = false;
            this.Browsebutton.Click += new System.EventHandler(this.Browsebutton_Click);
            // 
            // SaveButton
            // 
            this.SaveButton.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.SaveButton.Font = new System.Drawing.Font("Trebuchet MS", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SaveButton.Location = new System.Drawing.Point(293, 96);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(75, 34);
            this.SaveButton.TabIndex = 10;
            this.SaveButton.Text = "Save As";
            this.SaveButton.UseVisualStyleBackColor = false;
            this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // openFileDialog3
            // 
            this.openFileDialog3.Filter = "All Files|*.*|Word files (*.doc)|*.doc|Document (.docx)|*.doc|Portable Document F" +
    "ile (.pdf)|*.pdf";
            this.openFileDialog3.Title = "Select File";
            this.openFileDialog3.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog3_FileOk);
            // 
            // saveFileDialog3
            // 
            this.saveFileDialog3.CheckPathExists = false;
            this.saveFileDialog3.FileName = "Converted";
            this.saveFileDialog3.Filter = "Document (.docx)|*.docx|Word files (*.doc)|*.doc|Portable Document File (.pdf)|*." +
    "pdf|Text File (.txt)|*.txt";
            this.saveFileDialog3.Title = "Save As";
            this.saveFileDialog3.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog3_FileOk);
            // 
            // ConvertButton
            // 
            this.ConvertButton.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ConvertButton.Font = new System.Drawing.Font("Trebuchet MS", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ConvertButton.Location = new System.Drawing.Point(136, 162);
            this.ConvertButton.Name = "ConvertButton";
            this.ConvertButton.Size = new System.Drawing.Size(77, 33);
            this.ConvertButton.TabIndex = 11;
            this.ConvertButton.Text = "Convert";
            this.ConvertButton.UseVisualStyleBackColor = false;
            this.ConvertButton.Click += new System.EventHandler(this.ConvertButton_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(317, 196);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(72, 30);
            this.button1.TabIndex = 12;
            this.button1.Text = "Open File";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Trebuchet MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(26, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 16);
            this.label1.TabIndex = 13;
            this.label1.Text = "Open From:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Trebuchet MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(26, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 16);
            this.label2.TabIndex = 14;
            this.label2.Text = "Save To:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.menuStrip1.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.homeToolStripMenuItem,
            this.howToUseToolStripMenuItem,
            this.aboutAppToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.menuStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(401, 26);
            this.menuStrip1.TabIndex = 15;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // homeToolStripMenuItem
            // 
            this.homeToolStripMenuItem.Name = "homeToolStripMenuItem";
            this.homeToolStripMenuItem.Size = new System.Drawing.Size(51, 22);
            this.homeToolStripMenuItem.Text = "Home";
            // 
            // howToUseToolStripMenuItem
            // 
            this.howToUseToolStripMenuItem.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.howToUseToolStripMenuItem.Name = "howToUseToolStripMenuItem";
            this.howToUseToolStripMenuItem.Size = new System.Drawing.Size(86, 22);
            this.howToUseToolStripMenuItem.Text = "How To Use";
            this.howToUseToolStripMenuItem.Click += new System.EventHandler(this.howToUseToolStripMenuItem_Click);
            // 
            // aboutAppToolStripMenuItem
            // 
            this.aboutAppToolStripMenuItem.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.aboutAppToolStripMenuItem.Name = "aboutAppToolStripMenuItem";
            this.aboutAppToolStripMenuItem.Size = new System.Drawing.Size(53, 22);
            this.aboutAppToolStripMenuItem.Text = "About";
            this.aboutAppToolStripMenuItem.Click += new System.EventHandler(this.aboutAppToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(40, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(401, 254);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.ConvertButton);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.Browsebutton);
            this.Controls.Add(this.FileSave);
            this.Controls.Add(this.FileOpen);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Document Converter";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox FileOpen;
        private System.Windows.Forms.TextBox FileSave;
        private System.Windows.Forms.OpenFileDialog openFileDialog3;
        private System.Windows.Forms.SaveFileDialog saveFileDialog3;
        private System.Windows.Forms.Button ConvertButton;
        public System.Windows.Forms.Button Browsebutton;
        public System.Windows.Forms.Button SaveButton;
        public System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem howToUseToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutAppToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem homeToolStripMenuItem;
    }
}

