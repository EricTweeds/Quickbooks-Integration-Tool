namespace RedstoneQuickbooks
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
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("logo.ico")));
            this.locationDropdown = new System.Windows.Forms.ComboBox();
            this.dropdownTitle = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.importButton = new System.Windows.Forms.Button();
            this.consoleOutput = new System.Windows.Forms.TextBox();
            this.consoleTitle = new System.Windows.Forms.Label();
            this.uploadButton = new System.Windows.Forms.Button();
            this.doneButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // locationDropdown
            // 
            this.locationDropdown.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.locationDropdown.FormattingEnabled = true;
            this.locationDropdown.Items.AddRange(new object[] {
            "Redstone Retail",
            "Tawse Winery",
            "Farmers Market"});
            this.locationDropdown.Location = new System.Drawing.Point(16, 48);
            this.locationDropdown.Margin = new System.Windows.Forms.Padding(4);
            this.locationDropdown.Name = "locationDropdown";
            this.locationDropdown.Size = new System.Drawing.Size(160, 24);
            this.locationDropdown.TabIndex = 0;
            this.locationDropdown.SelectedIndexChanged += new System.EventHandler(this.locationDropdown_SelectedIndexChanged);
            // 
            // dropdownTitle
            // 
            this.dropdownTitle.AutoSize = true;
            this.dropdownTitle.Location = new System.Drawing.Point(13, 27);
            this.dropdownTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.dropdownTitle.Name = "dropdownTitle";
            this.dropdownTitle.Size = new System.Drawing.Size(111, 17);
            this.dropdownTitle.TabIndex = 1;
            this.dropdownTitle.Text = "Select Customer";
            this.dropdownTitle.Click += new System.EventHandler(this.dropdownTitle_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // importButton
            // 
            this.importButton.Location = new System.Drawing.Point(266, 44);
            this.importButton.Margin = new System.Windows.Forms.Padding(4);
            this.importButton.Name = "importButton";
            this.importButton.Size = new System.Drawing.Size(100, 28);
            this.importButton.TabIndex = 2;
            this.importButton.Text = "Import CSV";
            this.importButton.UseVisualStyleBackColor = true;
            this.importButton.Click += new System.EventHandler(this.importButton_Click);
            // 
            // consoleOutput
            // 
            this.consoleOutput.Location = new System.Drawing.Point(16, 150);
            this.consoleOutput.Margin = new System.Windows.Forms.Padding(4);
            this.consoleOutput.Multiline = true;
            this.consoleOutput.Name = "consoleOutput";
            this.consoleOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.consoleOutput.Size = new System.Drawing.Size(553, 138);
            this.consoleOutput.TabIndex = 3;
            // 
            // consoleTitle
            // 
            this.consoleTitle.AutoSize = true;
            this.consoleTitle.Location = new System.Drawing.Point(13, 129);
            this.consoleTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.consoleTitle.Name = "consoleTitle";
            this.consoleTitle.Size = new System.Drawing.Size(110, 17);
            this.consoleTitle.TabIndex = 4;
            this.consoleTitle.Text = "Console Output:";
            this.consoleTitle.Click += new System.EventHandler(this.consoleTitle_Click);
            // 
            // uploadButton
            // 
            this.uploadButton.Location = new System.Drawing.Point(595, 214);
            this.uploadButton.Margin = new System.Windows.Forms.Padding(4);
            this.uploadButton.Name = "uploadButton";
            this.uploadButton.Size = new System.Drawing.Size(100, 28);
            this.uploadButton.TabIndex = 5;
            this.uploadButton.Text = "Upload";
            this.uploadButton.UseVisualStyleBackColor = true;
            this.uploadButton.Click += new System.EventHandler(this.uploadButton_Click);
            // 
            // doneButton
            // 
            this.doneButton.Location = new System.Drawing.Point(595, 260);
            this.doneButton.Margin = new System.Windows.Forms.Padding(4);
            this.doneButton.Name = "doneButton";
            this.doneButton.Size = new System.Drawing.Size(100, 28);
            this.doneButton.TabIndex = 6;
            this.doneButton.Text = "Done";
            this.doneButton.UseVisualStyleBackColor = true;
            this.doneButton.Click += new System.EventHandler(this.doneButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(708, 321);
            this.Controls.Add(this.doneButton);
            this.Controls.Add(this.uploadButton);
            this.Controls.Add(this.consoleTitle);
            this.Controls.Add(this.consoleOutput);
            this.Controls.Add(this.importButton);
            this.Controls.Add(this.dropdownTitle);
            this.Controls.Add(this.locationDropdown);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Quickbooks Add Sales Reciept";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion



        private System.Windows.Forms.ComboBox locationDropdown;
        private System.Windows.Forms.Label dropdownTitle;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button importButton;
        private System.Windows.Forms.TextBox consoleOutput;
        private System.Windows.Forms.Label consoleTitle;
        private System.Windows.Forms.Button uploadButton;
        private System.Windows.Forms.Button doneButton;
    }
}

