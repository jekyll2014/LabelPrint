namespace FiscalLabelPrint
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.pictureBox_label = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBox_scale = new System.Windows.Forms.CheckBox();
            this.checkBox_toFile = new System.Windows.Forms.CheckBox();
            this.textBox_rangeTo = new System.Windows.Forms.TextBox();
            this.textBox_labelsName = new System.Windows.Forms.TextBox();
            this.textBox_templateName = new System.Windows.Forms.TextBox();
            this.textBox_saveFileName = new System.Windows.Forms.TextBox();
            this.textBox_rangeFrom = new System.Windows.Forms.TextBox();
            this.button_importTemplate = new System.Windows.Forms.Button();
            this.button_importLabels = new System.Windows.Forms.Button();
            this.button_printRange = new System.Windows.Forms.Button();
            this.button_printAll = new System.Windows.Forms.Button();
            this.button_printCurrent = new System.Windows.Forms.Button();
            this.dataGridView_labels = new System.Windows.Forms.DataGridView();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_label)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_labels)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.splitContainer2);
            this.splitContainer1.Panel1MinSize = 220;
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridView_labels);
            this.splitContainer1.Size = new System.Drawing.Size(384, 362);
            this.splitContainer1.SplitterDistance = 220;
            this.splitContainer1.TabIndex = 0;
            // 
            // splitContainer2
            // 
            this.splitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.AutoScroll = true;
            this.splitContainer2.Panel1.Controls.Add(this.pictureBox_label);
            this.splitContainer2.Panel1MinSize = 100;
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.label1);
            this.splitContainer2.Panel2.Controls.Add(this.checkBox_scale);
            this.splitContainer2.Panel2.Controls.Add(this.checkBox_toFile);
            this.splitContainer2.Panel2.Controls.Add(this.textBox_rangeTo);
            this.splitContainer2.Panel2.Controls.Add(this.textBox_labelsName);
            this.splitContainer2.Panel2.Controls.Add(this.textBox_templateName);
            this.splitContainer2.Panel2.Controls.Add(this.textBox_saveFileName);
            this.splitContainer2.Panel2.Controls.Add(this.textBox_rangeFrom);
            this.splitContainer2.Panel2.Controls.Add(this.button_importTemplate);
            this.splitContainer2.Panel2.Controls.Add(this.button_importLabels);
            this.splitContainer2.Panel2.Controls.Add(this.button_printRange);
            this.splitContainer2.Panel2.Controls.Add(this.button_printAll);
            this.splitContainer2.Panel2.Controls.Add(this.button_printCurrent);
            this.splitContainer2.Panel2MinSize = 220;
            this.splitContainer2.Size = new System.Drawing.Size(384, 220);
            this.splitContainer2.SplitterDistance = 160;
            this.splitContainer2.TabIndex = 0;
            // 
            // pictureBox_label
            // 
            this.pictureBox_label.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox_label.Location = new System.Drawing.Point(0, 0);
            this.pictureBox_label.Name = "pictureBox_label";
            this.pictureBox_label.Size = new System.Drawing.Size(156, 216);
            this.pictureBox_label.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox_label.TabIndex = 0;
            this.pictureBox_label.TabStop = false;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(149, 152);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(10, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "-";
            // 
            // checkBox_scale
            // 
            this.checkBox_scale.AutoSize = true;
            this.checkBox_scale.Location = new System.Drawing.Point(3, 68);
            this.checkBox_scale.Name = "checkBox_scale";
            this.checkBox_scale.Size = new System.Drawing.Size(106, 17);
            this.checkBox_scale.TabIndex = 2;
            this.checkBox_scale.Text = "Scale picture 1:1";
            this.checkBox_scale.UseVisualStyleBackColor = true;
            this.checkBox_scale.CheckedChanged += new System.EventHandler(this.checkBox_scale_CheckedChanged);
            // 
            // checkBox_toFile
            // 
            this.checkBox_toFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBox_toFile.AutoSize = true;
            this.checkBox_toFile.Location = new System.Drawing.Point(3, 177);
            this.checkBox_toFile.Name = "checkBox_toFile";
            this.checkBox_toFile.Size = new System.Drawing.Size(104, 17);
            this.checkBox_toFile.TabIndex = 2;
            this.checkBox_toFile.Text = "Print to .PNG file";
            this.checkBox_toFile.UseVisualStyleBackColor = true;
            this.checkBox_toFile.CheckedChanged += new System.EventHandler(this.checkBox_toFile_CheckedChanged);
            // 
            // textBox_rangeTo
            // 
            this.textBox_rangeTo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox_rangeTo.Location = new System.Drawing.Point(165, 149);
            this.textBox_rangeTo.Name = "textBox_rangeTo";
            this.textBox_rangeTo.Size = new System.Drawing.Size(44, 20);
            this.textBox_rangeTo.TabIndex = 1;
            this.textBox_rangeTo.Text = "00000";
            // 
            // textBox_labelsName
            // 
            this.textBox_labelsName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_labelsName.Enabled = false;
            this.textBox_labelsName.Location = new System.Drawing.Point(99, 41);
            this.textBox_labelsName.Name = "textBox_labelsName";
            this.textBox_labelsName.Size = new System.Drawing.Size(107, 20);
            this.textBox_labelsName.TabIndex = 1;
            // 
            // textBox_templateName
            // 
            this.textBox_templateName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_templateName.Enabled = false;
            this.textBox_templateName.Location = new System.Drawing.Point(99, 12);
            this.textBox_templateName.Name = "textBox_templateName";
            this.textBox_templateName.Size = new System.Drawing.Size(107, 20);
            this.textBox_templateName.TabIndex = 1;
            // 
            // textBox_saveFileName
            // 
            this.textBox_saveFileName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox_saveFileName.Enabled = false;
            this.textBox_saveFileName.Location = new System.Drawing.Point(113, 174);
            this.textBox_saveFileName.Name = "textBox_saveFileName";
            this.textBox_saveFileName.Size = new System.Drawing.Size(79, 20);
            this.textBox_saveFileName.TabIndex = 1;
            this.textBox_saveFileName.Text = "name";
            // 
            // textBox_rangeFrom
            // 
            this.textBox_rangeFrom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox_rangeFrom.Location = new System.Drawing.Point(99, 149);
            this.textBox_rangeFrom.Name = "textBox_rangeFrom";
            this.textBox_rangeFrom.Size = new System.Drawing.Size(44, 20);
            this.textBox_rangeFrom.TabIndex = 1;
            this.textBox_rangeFrom.Text = "00000";
            // 
            // button_importTemplate
            // 
            this.button_importTemplate.Location = new System.Drawing.Point(3, 10);
            this.button_importTemplate.Name = "button_importTemplate";
            this.button_importTemplate.Size = new System.Drawing.Size(90, 23);
            this.button_importTemplate.TabIndex = 0;
            this.button_importTemplate.Text = "Import template";
            this.button_importTemplate.UseVisualStyleBackColor = true;
            this.button_importTemplate.Click += new System.EventHandler(this.button_importTemplate_Click);
            // 
            // button_importLabels
            // 
            this.button_importLabels.Enabled = false;
            this.button_importLabels.Location = new System.Drawing.Point(3, 39);
            this.button_importLabels.Name = "button_importLabels";
            this.button_importLabels.Size = new System.Drawing.Size(90, 23);
            this.button_importLabels.TabIndex = 0;
            this.button_importLabels.Text = "Import labels";
            this.button_importLabels.UseVisualStyleBackColor = true;
            this.button_importLabels.Click += new System.EventHandler(this.button_importLabels_Click);
            // 
            // button_printRange
            // 
            this.button_printRange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_printRange.Enabled = false;
            this.button_printRange.Location = new System.Drawing.Point(3, 147);
            this.button_printRange.Name = "button_printRange";
            this.button_printRange.Size = new System.Drawing.Size(90, 23);
            this.button_printRange.TabIndex = 0;
            this.button_printRange.Text = "Print range";
            this.button_printRange.UseVisualStyleBackColor = true;
            this.button_printRange.Click += new System.EventHandler(this.button_printRange_Click);
            // 
            // button_printAll
            // 
            this.button_printAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_printAll.Enabled = false;
            this.button_printAll.Location = new System.Drawing.Point(3, 118);
            this.button_printAll.Name = "button_printAll";
            this.button_printAll.Size = new System.Drawing.Size(90, 23);
            this.button_printAll.TabIndex = 0;
            this.button_printAll.Text = "Print all";
            this.button_printAll.UseVisualStyleBackColor = true;
            this.button_printAll.Click += new System.EventHandler(this.button_printAll_Click);
            // 
            // button_printCurrent
            // 
            this.button_printCurrent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_printCurrent.Enabled = false;
            this.button_printCurrent.Location = new System.Drawing.Point(3, 89);
            this.button_printCurrent.Name = "button_printCurrent";
            this.button_printCurrent.Size = new System.Drawing.Size(90, 23);
            this.button_printCurrent.TabIndex = 0;
            this.button_printCurrent.Text = "Print current";
            this.button_printCurrent.UseVisualStyleBackColor = true;
            this.button_printCurrent.Click += new System.EventHandler(this.button_printCurrent_Click);
            // 
            // dataGridView_labels
            // 
            this.dataGridView_labels.AllowUserToAddRows = false;
            this.dataGridView_labels.AllowUserToDeleteRows = false;
            this.dataGridView_labels.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_labels.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_labels.Location = new System.Drawing.Point(0, 0);
            this.dataGridView_labels.MultiSelect = false;
            this.dataGridView_labels.Name = "dataGridView_labels";
            this.dataGridView_labels.ReadOnly = true;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.Format = "N0";
            dataGridViewCellStyle1.NullValue = null;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_labels.RowHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView_labels.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_labels.Size = new System.Drawing.Size(380, 134);
            this.dataGridView_labels.TabIndex = 0;
            this.dataGridView_labels.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_labels_CellClick);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // printDialog1
            // 
            this.printDialog1.UseEXDialog = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 362);
            this.Controls.Add(this.splitContainer1);
            this.MinimumSize = new System.Drawing.Size(400, 400);
            this.Name = "Form1";
            this.Text = "FiscalLabelPrint";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_label)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_labels)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.PictureBox pictureBox_label;
        private System.Windows.Forms.Button button_printCurrent;
        private System.Windows.Forms.DataGridView dataGridView_labels;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textBox_rangeTo;
        private System.Windows.Forms.TextBox textBox_rangeFrom;
        private System.Windows.Forms.Button button_printRange;
        private System.Windows.Forms.Button button_printAll;
        private System.Windows.Forms.Button button_importLabels;
        private System.Windows.Forms.CheckBox checkBox_toFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_saveFileName;
        private System.Windows.Forms.Button button_importTemplate;
        private System.Windows.Forms.PrintDialog printDialog1;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.CheckBox checkBox_scale;
        private System.Windows.Forms.TextBox textBox_labelsName;
        private System.Windows.Forms.TextBox textBox_templateName;
    }
}

