namespace DCH_OpenBalance
{
    partial class DCH_OpenBalance
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
            this.components = new System.ComponentModel.Container();
            this.button_Choose = new System.Windows.Forms.Button();
            this.dataGridView_Display = new System.Windows.Forms.DataGridView();
            this.button_Save = new System.Windows.Forms.Button();
            this.button_Display = new System.Windows.Forms.Button();
            this.button_ExportExcel = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.textBox_Path = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button_Browse = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_Name = new System.Windows.Forms.TextBox();
            this.button_Validate = new System.Windows.Forms.Button();
            this.radioButton_Vaidate = new System.Windows.Forms.RadioButton();
            this.radioButton_success = new System.Windows.Forms.RadioButton();
            this.button_Complete = new System.Windows.Forms.Button();
            this.tmrWork = new System.Windows.Forms.Timer(this.components);
            this.button_ExportExpireDate = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Display)).BeginInit();
            this.SuspendLayout();
            // 
            // button_Choose
            // 
            this.button_Choose.Location = new System.Drawing.Point(15, 676);
            this.button_Choose.Name = "button_Choose";
            this.button_Choose.Size = new System.Drawing.Size(158, 23);
            this.button_Choose.TabIndex = 0;
            this.button_Choose.Text = "Choose and Read File Excel";
            this.button_Choose.UseVisualStyleBackColor = true;
            this.button_Choose.Click += new System.EventHandler(this.Button_Choose_Click);
            // 
            // dataGridView_Display
            // 
            this.dataGridView_Display.AllowUserToAddRows = false;
            this.dataGridView_Display.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView_Display.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_Display.Location = new System.Drawing.Point(3, 135);
            this.dataGridView_Display.Name = "dataGridView_Display";
            this.dataGridView_Display.Size = new System.Drawing.Size(1128, 514);
            this.dataGridView_Display.TabIndex = 1;
            // 
            // button_Save
            // 
            this.button_Save.Location = new System.Drawing.Point(190, 676);
            this.button_Save.Name = "button_Save";
            this.button_Save.Size = new System.Drawing.Size(102, 23);
            this.button_Save.TabIndex = 2;
            this.button_Save.Text = "Save";
            this.button_Save.UseVisualStyleBackColor = true;
            this.button_Save.Click += new System.EventHandler(this.Button_Save_Click);
            // 
            // button_Display
            // 
            this.button_Display.Location = new System.Drawing.Point(471, 676);
            this.button_Display.Name = "button_Display";
            this.button_Display.Size = new System.Drawing.Size(148, 23);
            this.button_Display.TabIndex = 3;
            this.button_Display.Text = "Display Error";
            this.button_Display.UseVisualStyleBackColor = true;
            this.button_Display.Click += new System.EventHandler(this.Button_Display_Click);
            // 
            // button_ExportExcel
            // 
            this.button_ExportExcel.Location = new System.Drawing.Point(804, 676);
            this.button_ExportExcel.Name = "button_ExportExcel";
            this.button_ExportExcel.Size = new System.Drawing.Size(148, 23);
            this.button_ExportExcel.TabIndex = 4;
            this.button_ExportExcel.Text = "Export Excel";
            this.button_ExportExcel.UseVisualStyleBackColor = true;
            this.button_ExportExcel.Click += new System.EventHandler(this.Button_ExportExcel_Click);
            // 
            // textBox_Path
            // 
            this.textBox_Path.Enabled = false;
            this.textBox_Path.Location = new System.Drawing.Point(72, 12);
            this.textBox_Path.Name = "textBox_Path";
            this.textBox_Path.Size = new System.Drawing.Size(402, 20);
            this.textBox_Path.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "File Path :";
            // 
            // button_Browse
            // 
            this.button_Browse.Location = new System.Drawing.Point(480, 10);
            this.button_Browse.Name = "button_Browse";
            this.button_Browse.Size = new System.Drawing.Size(42, 22);
            this.button_Browse.TabIndex = 7;
            this.button_Browse.Text = "...";
            this.button_Browse.UseVisualStyleBackColor = true;
            this.button_Browse.Click += new System.EventHandler(this.Button_Browse_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "File Name :";
            // 
            // textBox_Name
            // 
            this.textBox_Name.Location = new System.Drawing.Point(72, 38);
            this.textBox_Name.Name = "textBox_Name";
            this.textBox_Name.Size = new System.Drawing.Size(402, 20);
            this.textBox_Name.TabIndex = 8;
            this.textBox_Name.Text = "DCH_OpenBalance";
            // 
            // button_Validate
            // 
            this.button_Validate.Location = new System.Drawing.Point(309, 676);
            this.button_Validate.Name = "button_Validate";
            this.button_Validate.Size = new System.Drawing.Size(148, 23);
            this.button_Validate.TabIndex = 10;
            this.button_Validate.Text = "Validate Data";
            this.button_Validate.UseVisualStyleBackColor = true;
            this.button_Validate.Click += new System.EventHandler(this.Button_Validate_Click);
            // 
            // radioButton_Vaidate
            // 
            this.radioButton_Vaidate.AutoSize = true;
            this.radioButton_Vaidate.Checked = true;
            this.radioButton_Vaidate.Location = new System.Drawing.Point(15, 64);
            this.radioButton_Vaidate.Name = "radioButton_Vaidate";
            this.radioButton_Vaidate.Size = new System.Drawing.Size(174, 17);
            this.radioButton_Vaidate.TabIndex = 12;
            this.radioButton_Vaidate.TabStop = true;
            this.radioButton_Vaidate.Text = "Export data not correct to Excel";
            this.radioButton_Vaidate.UseVisualStyleBackColor = true;
            // 
            // radioButton_success
            // 
            this.radioButton_success.AutoSize = true;
            this.radioButton_success.Location = new System.Drawing.Point(195, 64);
            this.radioButton_success.Name = "radioButton_success";
            this.radioButton_success.Size = new System.Drawing.Size(156, 17);
            this.radioButton_success.TabIndex = 13;
            this.radioButton_success.Text = "Export data correct to Excel";
            this.radioButton_success.UseVisualStyleBackColor = true;
            // 
            // button_Complete
            // 
            this.button_Complete.Location = new System.Drawing.Point(637, 676);
            this.button_Complete.Name = "button_Complete";
            this.button_Complete.Size = new System.Drawing.Size(148, 23);
            this.button_Complete.TabIndex = 14;
            this.button_Complete.Text = "Display comlete";
            this.button_Complete.UseVisualStyleBackColor = true;
            this.button_Complete.Click += new System.EventHandler(this.Button_Complete_Click);
            // 
            // button_ExportExpireDate
            // 
            this.button_ExportExpireDate.Location = new System.Drawing.Point(958, 676);
            this.button_ExportExpireDate.Name = "button_ExportExpireDate";
            this.button_ExportExpireDate.Size = new System.Drawing.Size(173, 23);
            this.button_ExportExpireDate.TabIndex = 15;
            this.button_ExportExpireDate.Text = "Export Excel Expire Date";
            this.button_ExportExpireDate.UseVisualStyleBackColor = true;
            this.button_ExportExpireDate.Click += new System.EventHandler(this.Button_ExportExpireDate_Click);
            // 
            // DCH_OpenBalance
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1143, 711);
            this.Controls.Add(this.button_ExportExpireDate);
            this.Controls.Add(this.button_Complete);
            this.Controls.Add(this.radioButton_success);
            this.Controls.Add(this.radioButton_Vaidate);
            this.Controls.Add(this.button_Validate);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox_Name);
            this.Controls.Add(this.button_Browse);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox_Path);
            this.Controls.Add(this.button_ExportExcel);
            this.Controls.Add(this.button_Display);
            this.Controls.Add(this.button_Save);
            this.Controls.Add(this.dataGridView_Display);
            this.Controls.Add(this.button_Choose);
            this.Name = "DCH_OpenBalance";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "DCH OpenBalance";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.DCH_OpenBalance_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Display)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_Choose;
        private System.Windows.Forms.DataGridView dataGridView_Display;
        private System.Windows.Forms.Button button_Save;
        private System.Windows.Forms.Button button_Display;
        private System.Windows.Forms.Button button_ExportExcel;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox textBox_Path;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_Browse;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_Name;
        private System.Windows.Forms.Button button_Validate;
        private System.Windows.Forms.RadioButton radioButton_Vaidate;
        private System.Windows.Forms.RadioButton radioButton_success;
        private System.Windows.Forms.Button button_Complete;
        private System.Windows.Forms.Timer tmrWork;
        private System.Windows.Forms.Button button_ExportExpireDate;
    }
}

