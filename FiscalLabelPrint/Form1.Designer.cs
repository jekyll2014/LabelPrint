namespace LabelPrint
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
            LabelBmp.Dispose();
            objectBmp.Dispose();
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.splitContainer_horizontal = new System.Windows.Forms.SplitContainer();
            this.splitContainer_vertical = new System.Windows.Forms.SplitContainer();
            this.pictureBox_label = new System.Windows.Forms.PictureBox();
            this.textBox_mX = new System.Windows.Forms.TextBox();
            this.textBox_mY = new System.Windows.Forms.TextBox();
            this.checkBox_scale = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage_print = new System.Windows.Forms.TabPage();
            this.button_printRange = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.button_printCurrent = new System.Windows.Forms.Button();
            this.button_printAll = new System.Windows.Forms.Button();
            this.textBox_rangeFrom = new System.Windows.Forms.TextBox();
            this.checkBox_toFile = new System.Windows.Forms.CheckBox();
            this.textBox_saveFileName = new System.Windows.Forms.TextBox();
            this.textBox_rangeTo = new System.Windows.Forms.TextBox();
            this.tabPage_edit = new System.Windows.Forms.TabPage();
            this.listBox_objects = new System.Windows.Forms.ListBox();
            this.label_content = new System.Windows.Forms.Label();
            this.label_fontSize = new System.Windows.Forms.Label();
            this.label_rotate = new System.Windows.Forms.Label();
            this.label_posY = new System.Windows.Forms.Label();
            this.label_height = new System.Windows.Forms.Label();
            this.label_width = new System.Windows.Forms.Label();
            this.label_posX = new System.Windows.Forms.Label();
            this.label_fontName = new System.Windows.Forms.Label();
            this.label_fontStyle = new System.Windows.Forms.Label();
            this.label_objectColor = new System.Windows.Forms.Label();
            this.label_backgroundColor = new System.Windows.Forms.Label();
            this.checkBox_fill = new System.Windows.Forms.CheckBox();
            this.textBox_height = new System.Windows.Forms.TextBox();
            this.textBox_fontSize = new System.Windows.Forms.TextBox();
            this.textBox_rotate = new System.Windows.Forms.TextBox();
            this.textBox_width = new System.Windows.Forms.TextBox();
            this.textBox_posY = new System.Windows.Forms.TextBox();
            this.textBox_content = new System.Windows.Forms.TextBox();
            this.textBox_posX = new System.Windows.Forms.TextBox();
            this.comboBox_objectColor = new System.Windows.Forms.ComboBox();
            this.comboBox_fontStyle = new System.Windows.Forms.ComboBox();
            this.comboBox_object = new System.Windows.Forms.ComboBox();
            this.comboBox_fontName = new System.Windows.Forms.ComboBox();
            this.comboBox_backgroundColor = new System.Windows.Forms.ComboBox();
            this.button_clone = new System.Windows.Forms.Button();
            this.button_saveLabel = new System.Windows.Forms.Button();
            this.button_saveTemplate = new System.Windows.Forms.Button();
            this.button_delete = new System.Windows.Forms.Button();
            this.button_down = new System.Windows.Forms.Button();
            this.button_up = new System.Windows.Forms.Button();
            this.button_apply = new System.Windows.Forms.Button();
            this.tabPage_group = new System.Windows.Forms.TabPage();
            this.listBox_objectsMulti = new System.Windows.Forms.ListBox();
            this.button_deleteGroup = new System.Windows.Forms.Button();
            this.button_moveLeft = new System.Windows.Forms.Button();
            this.button_moveDown = new System.Windows.Forms.Button();
            this.button_moveUp = new System.Windows.Forms.Button();
            this.textBox_move = new System.Windows.Forms.TextBox();
            this.button_moveRight = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox_dpi = new System.Windows.Forms.TextBox();
            this.checkBox_columnNames = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_labelsName = new System.Windows.Forms.TextBox();
            this.comboBox_units = new System.Windows.Forms.ComboBox();
            this.textBox_templateName = new System.Windows.Forms.TextBox();
            this.button_importTemplate = new System.Windows.Forms.Button();
            this.button_importLabels = new System.Windows.Forms.Button();
            this.comboBox_encoding = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dataGridView_labels = new System.Windows.Forms.DataGridView();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.SaveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer_horizontal)).BeginInit();
            this.splitContainer_horizontal.Panel1.SuspendLayout();
            this.splitContainer_horizontal.Panel2.SuspendLayout();
            this.splitContainer_horizontal.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer_vertical)).BeginInit();
            this.splitContainer_vertical.Panel1.SuspendLayout();
            this.splitContainer_vertical.Panel2.SuspendLayout();
            this.splitContainer_vertical.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_label)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage_print.SuspendLayout();
            this.tabPage_edit.SuspendLayout();
            this.tabPage_group.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_labels)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer_horizontal
            // 
            this.splitContainer_horizontal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.splitContainer_horizontal.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer_horizontal.Location = new System.Drawing.Point(0, 0);
            this.splitContainer_horizontal.Name = "splitContainer_horizontal";
            this.splitContainer_horizontal.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer_horizontal.Panel1
            // 
            this.splitContainer_horizontal.Panel1.Controls.Add(this.splitContainer_vertical);
            this.splitContainer_horizontal.Panel1MinSize = 460;
            // 
            // splitContainer_horizontal.Panel2
            // 
            this.splitContainer_horizontal.Panel2.Controls.Add(this.dataGridView_labels);
            this.splitContainer_horizontal.Panel2MinSize = 100;
            this.splitContainer_horizontal.Size = new System.Drawing.Size(958, 654);
            this.splitContainer_horizontal.SplitterDistance = 488;
            this.splitContainer_horizontal.TabIndex = 0;
            // 
            // splitContainer_vertical
            // 
            this.splitContainer_vertical.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.splitContainer_vertical.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer_vertical.Location = new System.Drawing.Point(0, 0);
            this.splitContainer_vertical.Name = "splitContainer_vertical";
            // 
            // splitContainer_vertical.Panel1
            // 
            this.splitContainer_vertical.Panel1.AutoScroll = true;
            this.splitContainer_vertical.Panel1.Controls.Add(this.pictureBox_label);
            this.splitContainer_vertical.Panel1.Controls.Add(this.textBox_mX);
            this.splitContainer_vertical.Panel1.Controls.Add(this.textBox_mY);
            this.splitContainer_vertical.Panel1.Controls.Add(this.checkBox_scale);
            this.splitContainer_vertical.Panel1MinSize = 110;
            // 
            // splitContainer_vertical.Panel2
            // 
            this.splitContainer_vertical.Panel2.Controls.Add(this.label5);
            this.splitContainer_vertical.Panel2.Controls.Add(this.tabControl1);
            this.splitContainer_vertical.Panel2.Controls.Add(this.label4);
            this.splitContainer_vertical.Panel2.Controls.Add(this.textBox_dpi);
            this.splitContainer_vertical.Panel2.Controls.Add(this.checkBox_columnNames);
            this.splitContainer_vertical.Panel2.Controls.Add(this.label3);
            this.splitContainer_vertical.Panel2.Controls.Add(this.textBox_labelsName);
            this.splitContainer_vertical.Panel2.Controls.Add(this.comboBox_units);
            this.splitContainer_vertical.Panel2.Controls.Add(this.textBox_templateName);
            this.splitContainer_vertical.Panel2.Controls.Add(this.button_importTemplate);
            this.splitContainer_vertical.Panel2.Controls.Add(this.button_importLabels);
            this.splitContainer_vertical.Panel2.Controls.Add(this.comboBox_encoding);
            this.splitContainer_vertical.Panel2.Controls.Add(this.label2);
            this.splitContainer_vertical.Panel2MinSize = 350;
            this.splitContainer_vertical.Size = new System.Drawing.Size(958, 488);
            this.splitContainer_vertical.SplitterDistance = 296;
            this.splitContainer_vertical.TabIndex = 0;
            // 
            // pictureBox_label
            // 
            this.pictureBox_label.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox_label.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox_label.Location = new System.Drawing.Point(0, 0);
            this.pictureBox_label.Name = "pictureBox_label";
            this.pictureBox_label.Size = new System.Drawing.Size(289, 457);
            this.pictureBox_label.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox_label.TabIndex = 0;
            this.pictureBox_label.TabStop = false;
            this.pictureBox_label.MouseMove += new System.Windows.Forms.MouseEventHandler(this.PictureBox_label_MouseMove);
            // 
            // textBox_mX
            // 
            this.textBox_mX.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox_mX.Enabled = false;
            this.textBox_mX.Location = new System.Drawing.Point(3, 461);
            this.textBox_mX.MinimumSize = new System.Drawing.Size(60, 4);
            this.textBox_mX.Name = "textBox_mX";
            this.textBox_mX.Size = new System.Drawing.Size(67, 20);
            this.textBox_mX.TabIndex = 3;
            // 
            // textBox_mY
            // 
            this.textBox_mY.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox_mY.Enabled = false;
            this.textBox_mY.Location = new System.Drawing.Point(79, 461);
            this.textBox_mY.MinimumSize = new System.Drawing.Size(60, 4);
            this.textBox_mY.Name = "textBox_mY";
            this.textBox_mY.Size = new System.Drawing.Size(67, 20);
            this.textBox_mY.TabIndex = 3;
            // 
            // checkBox_scale
            // 
            this.checkBox_scale.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBox_scale.AutoSize = true;
            this.checkBox_scale.Location = new System.Drawing.Point(152, 463);
            this.checkBox_scale.Name = "checkBox_scale";
            this.checkBox_scale.Size = new System.Drawing.Size(106, 17);
            this.checkBox_scale.TabIndex = 2;
            this.checkBox_scale.Text = "Scale picture 1:1";
            this.checkBox_scale.UseVisualStyleBackColor = true;
            this.checkBox_scale.CheckedChanged += new System.EventHandler(this.CheckBox_scale_CheckedChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(116, 98);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(21, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "dpi";
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage_print);
            this.tabControl1.Controls.Add(this.tabPage_edit);
            this.tabControl1.Controls.Add(this.tabPage_group);
            this.tabControl1.Location = new System.Drawing.Point(3, 121);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(641, 360);
            this.tabControl1.TabIndex = 4;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.TabControl1_SelectedIndexChanged);
            // 
            // tabPage_print
            // 
            this.tabPage_print.Controls.Add(this.button_printRange);
            this.tabPage_print.Controls.Add(this.label1);
            this.tabPage_print.Controls.Add(this.button_printCurrent);
            this.tabPage_print.Controls.Add(this.button_printAll);
            this.tabPage_print.Controls.Add(this.textBox_rangeFrom);
            this.tabPage_print.Controls.Add(this.checkBox_toFile);
            this.tabPage_print.Controls.Add(this.textBox_saveFileName);
            this.tabPage_print.Controls.Add(this.textBox_rangeTo);
            this.tabPage_print.Location = new System.Drawing.Point(4, 22);
            this.tabPage_print.Name = "tabPage_print";
            this.tabPage_print.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_print.Size = new System.Drawing.Size(633, 334);
            this.tabPage_print.TabIndex = 0;
            this.tabPage_print.Text = "Print labels";
            this.tabPage_print.UseVisualStyleBackColor = true;
            // 
            // button_printRange
            // 
            this.button_printRange.Enabled = false;
            this.button_printRange.Location = new System.Drawing.Point(6, 87);
            this.button_printRange.Name = "button_printRange";
            this.button_printRange.Size = new System.Drawing.Size(90, 23);
            this.button_printRange.TabIndex = 0;
            this.button_printRange.Text = "Print range";
            this.button_printRange.UseVisualStyleBackColor = true;
            this.button_printRange.Click += new System.EventHandler(this.Button_printRange_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(152, 92);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(10, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "-";
            // 
            // button_printCurrent
            // 
            this.button_printCurrent.Enabled = false;
            this.button_printCurrent.Location = new System.Drawing.Point(6, 29);
            this.button_printCurrent.Name = "button_printCurrent";
            this.button_printCurrent.Size = new System.Drawing.Size(90, 23);
            this.button_printCurrent.TabIndex = 0;
            this.button_printCurrent.Text = "Print current";
            this.button_printCurrent.UseVisualStyleBackColor = true;
            this.button_printCurrent.Click += new System.EventHandler(this.Button_printCurrent_Click);
            // 
            // button_printAll
            // 
            this.button_printAll.Enabled = false;
            this.button_printAll.Location = new System.Drawing.Point(6, 58);
            this.button_printAll.Name = "button_printAll";
            this.button_printAll.Size = new System.Drawing.Size(90, 23);
            this.button_printAll.TabIndex = 0;
            this.button_printAll.Text = "Print all";
            this.button_printAll.UseVisualStyleBackColor = true;
            this.button_printAll.Click += new System.EventHandler(this.Button_printAll_Click);
            // 
            // textBox_rangeFrom
            // 
            this.textBox_rangeFrom.Location = new System.Drawing.Point(102, 89);
            this.textBox_rangeFrom.Name = "textBox_rangeFrom";
            this.textBox_rangeFrom.Size = new System.Drawing.Size(44, 20);
            this.textBox_rangeFrom.TabIndex = 1;
            this.textBox_rangeFrom.Text = "1";
            this.textBox_rangeFrom.Leave += new System.EventHandler(this.textBox_rangeFrom_Leave);
            // 
            // checkBox_toFile
            // 
            this.checkBox_toFile.AutoSize = true;
            this.checkBox_toFile.Location = new System.Drawing.Point(6, 6);
            this.checkBox_toFile.Name = "checkBox_toFile";
            this.checkBox_toFile.Size = new System.Drawing.Size(104, 17);
            this.checkBox_toFile.TabIndex = 2;
            this.checkBox_toFile.Text = "Print to .PNG file";
            this.checkBox_toFile.UseVisualStyleBackColor = true;
            this.checkBox_toFile.CheckedChanged += new System.EventHandler(this.CheckBox_toFile_CheckedChanged);
            // 
            // textBox_saveFileName
            // 
            this.textBox_saveFileName.Enabled = false;
            this.textBox_saveFileName.Location = new System.Drawing.Point(116, 3);
            this.textBox_saveFileName.Name = "textBox_saveFileName";
            this.textBox_saveFileName.Size = new System.Drawing.Size(79, 20);
            this.textBox_saveFileName.TabIndex = 1;
            this.textBox_saveFileName.Text = "name";
            // 
            // textBox_rangeTo
            // 
            this.textBox_rangeTo.Location = new System.Drawing.Point(168, 89);
            this.textBox_rangeTo.Name = "textBox_rangeTo";
            this.textBox_rangeTo.Size = new System.Drawing.Size(44, 20);
            this.textBox_rangeTo.TabIndex = 1;
            this.textBox_rangeTo.Text = "1";
            this.textBox_rangeTo.Leave += new System.EventHandler(this.textBox_rangeTo_Leave);
            // 
            // tabPage_edit
            // 
            this.tabPage_edit.AutoScroll = true;
            this.tabPage_edit.Controls.Add(this.listBox_objects);
            this.tabPage_edit.Controls.Add(this.label_content);
            this.tabPage_edit.Controls.Add(this.label_fontSize);
            this.tabPage_edit.Controls.Add(this.label_rotate);
            this.tabPage_edit.Controls.Add(this.label_posY);
            this.tabPage_edit.Controls.Add(this.label_height);
            this.tabPage_edit.Controls.Add(this.label_width);
            this.tabPage_edit.Controls.Add(this.label_posX);
            this.tabPage_edit.Controls.Add(this.label_fontName);
            this.tabPage_edit.Controls.Add(this.label_fontStyle);
            this.tabPage_edit.Controls.Add(this.label_objectColor);
            this.tabPage_edit.Controls.Add(this.label_backgroundColor);
            this.tabPage_edit.Controls.Add(this.checkBox_fill);
            this.tabPage_edit.Controls.Add(this.textBox_height);
            this.tabPage_edit.Controls.Add(this.textBox_fontSize);
            this.tabPage_edit.Controls.Add(this.textBox_rotate);
            this.tabPage_edit.Controls.Add(this.textBox_width);
            this.tabPage_edit.Controls.Add(this.textBox_posY);
            this.tabPage_edit.Controls.Add(this.textBox_content);
            this.tabPage_edit.Controls.Add(this.textBox_posX);
            this.tabPage_edit.Controls.Add(this.comboBox_objectColor);
            this.tabPage_edit.Controls.Add(this.comboBox_fontStyle);
            this.tabPage_edit.Controls.Add(this.comboBox_object);
            this.tabPage_edit.Controls.Add(this.comboBox_fontName);
            this.tabPage_edit.Controls.Add(this.comboBox_backgroundColor);
            this.tabPage_edit.Controls.Add(this.button_clone);
            this.tabPage_edit.Controls.Add(this.button_saveLabel);
            this.tabPage_edit.Controls.Add(this.button_saveTemplate);
            this.tabPage_edit.Controls.Add(this.button_delete);
            this.tabPage_edit.Controls.Add(this.button_down);
            this.tabPage_edit.Controls.Add(this.button_up);
            this.tabPage_edit.Controls.Add(this.button_apply);
            this.tabPage_edit.Location = new System.Drawing.Point(4, 22);
            this.tabPage_edit.Name = "tabPage_edit";
            this.tabPage_edit.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_edit.Size = new System.Drawing.Size(633, 334);
            this.tabPage_edit.TabIndex = 1;
            this.tabPage_edit.Text = "Edit template";
            this.tabPage_edit.UseVisualStyleBackColor = true;
            // 
            // listBox_objects
            // 
            this.listBox_objects.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.listBox_objects.FormattingEnabled = true;
            this.listBox_objects.Location = new System.Drawing.Point(7, 36);
            this.listBox_objects.Name = "listBox_objects";
            this.listBox_objects.Size = new System.Drawing.Size(95, 199);
            this.listBox_objects.TabIndex = 6;
            this.listBox_objects.SelectedIndexChanged += new System.EventHandler(this.ListBox_objects_SelectedIndexChanged);
            // 
            // label_content
            // 
            this.label_content.AutoSize = true;
            this.label_content.Location = new System.Drawing.Point(108, 257);
            this.label_content.Name = "label_content";
            this.label_content.Size = new System.Drawing.Size(80, 13);
            this.label_content.TabIndex = 5;
            this.label_content.Text = "Default content";
            // 
            // label_fontSize
            // 
            this.label_fontSize.AutoSize = true;
            this.label_fontSize.Location = new System.Drawing.Point(181, 217);
            this.label_fontSize.Name = "label_fontSize";
            this.label_fontSize.Size = new System.Drawing.Size(49, 13);
            this.label_fontSize.TabIndex = 5;
            this.label_fontSize.Text = "Font size";
            // 
            // label_rotate
            // 
            this.label_rotate.AutoSize = true;
            this.label_rotate.Location = new System.Drawing.Point(108, 218);
            this.label_rotate.Name = "label_rotate";
            this.label_rotate.Size = new System.Drawing.Size(39, 13);
            this.label_rotate.TabIndex = 5;
            this.label_rotate.Text = "Rotate";
            // 
            // label_posY
            // 
            this.label_posY.AutoSize = true;
            this.label_posY.Location = new System.Drawing.Point(108, 101);
            this.label_posY.Name = "label_posY";
            this.label_posY.Size = new System.Drawing.Size(31, 13);
            this.label_posY.TabIndex = 5;
            this.label_posY.Text = "posY";
            // 
            // label_height
            // 
            this.label_height.AutoSize = true;
            this.label_height.Location = new System.Drawing.Point(111, 179);
            this.label_height.Name = "label_height";
            this.label_height.Size = new System.Drawing.Size(38, 13);
            this.label_height.TabIndex = 5;
            this.label_height.Text = "Height";
            // 
            // label_width
            // 
            this.label_width.AutoSize = true;
            this.label_width.Location = new System.Drawing.Point(108, 140);
            this.label_width.Name = "label_width";
            this.label_width.Size = new System.Drawing.Size(35, 13);
            this.label_width.TabIndex = 5;
            this.label_width.Text = "Width";
            // 
            // label_posX
            // 
            this.label_posX.AutoSize = true;
            this.label_posX.Location = new System.Drawing.Point(108, 62);
            this.label_posX.Name = "label_posX";
            this.label_posX.Size = new System.Drawing.Size(31, 13);
            this.label_posX.TabIndex = 5;
            this.label_posX.Text = "posX";
            // 
            // label_fontName
            // 
            this.label_fontName.AutoSize = true;
            this.label_fontName.Location = new System.Drawing.Point(181, 139);
            this.label_fontName.Name = "label_fontName";
            this.label_fontName.Size = new System.Drawing.Size(57, 13);
            this.label_fontName.TabIndex = 5;
            this.label_fontName.Text = "Font name";
            // 
            // label_fontStyle
            // 
            this.label_fontStyle.AutoSize = true;
            this.label_fontStyle.Location = new System.Drawing.Point(181, 178);
            this.label_fontStyle.Name = "label_fontStyle";
            this.label_fontStyle.Size = new System.Drawing.Size(52, 13);
            this.label_fontStyle.TabIndex = 5;
            this.label_fontStyle.Text = "Font style";
            // 
            // label_objectColor
            // 
            this.label_objectColor.AutoSize = true;
            this.label_objectColor.Location = new System.Drawing.Point(181, 100);
            this.label_objectColor.Name = "label_objectColor";
            this.label_objectColor.Size = new System.Drawing.Size(64, 13);
            this.label_objectColor.TabIndex = 5;
            this.label_objectColor.Text = "Object color";
            // 
            // label_backgroundColor
            // 
            this.label_backgroundColor.AutoSize = true;
            this.label_backgroundColor.Location = new System.Drawing.Point(181, 61);
            this.label_backgroundColor.Name = "label_backgroundColor";
            this.label_backgroundColor.Size = new System.Drawing.Size(91, 13);
            this.label_backgroundColor.TabIndex = 5;
            this.label_backgroundColor.Text = "Background color";
            // 
            // checkBox_fill
            // 
            this.checkBox_fill.AutoSize = true;
            this.checkBox_fill.Location = new System.Drawing.Point(190, 37);
            this.checkBox_fill.Name = "checkBox_fill";
            this.checkBox_fill.Size = new System.Drawing.Size(83, 17);
            this.checkBox_fill.TabIndex = 4;
            this.checkBox_fill.Text = "Transparent";
            this.checkBox_fill.UseVisualStyleBackColor = true;
            // 
            // textBox_height
            // 
            this.textBox_height.Location = new System.Drawing.Point(108, 195);
            this.textBox_height.MinimumSize = new System.Drawing.Size(60, 4);
            this.textBox_height.Name = "textBox_height";
            this.textBox_height.Size = new System.Drawing.Size(67, 20);
            this.textBox_height.TabIndex = 3;
            this.textBox_height.Leave += new System.EventHandler(this.TextBox_height_Leave);
            // 
            // textBox_fontSize
            // 
            this.textBox_fontSize.Location = new System.Drawing.Point(181, 233);
            this.textBox_fontSize.MinimumSize = new System.Drawing.Size(135, 4);
            this.textBox_fontSize.Name = "textBox_fontSize";
            this.textBox_fontSize.Size = new System.Drawing.Size(135, 20);
            this.textBox_fontSize.TabIndex = 3;
            // 
            // textBox_rotate
            // 
            this.textBox_rotate.Location = new System.Drawing.Point(108, 233);
            this.textBox_rotate.MinimumSize = new System.Drawing.Size(60, 4);
            this.textBox_rotate.Name = "textBox_rotate";
            this.textBox_rotate.Size = new System.Drawing.Size(67, 20);
            this.textBox_rotate.TabIndex = 3;
            this.textBox_rotate.Leave += new System.EventHandler(this.TextBox_rotate_Leave);
            // 
            // textBox_width
            // 
            this.textBox_width.Location = new System.Drawing.Point(108, 156);
            this.textBox_width.MinimumSize = new System.Drawing.Size(60, 4);
            this.textBox_width.Name = "textBox_width";
            this.textBox_width.Size = new System.Drawing.Size(67, 20);
            this.textBox_width.TabIndex = 3;
            this.textBox_width.Leave += new System.EventHandler(this.TextBox_width_Leave);
            // 
            // textBox_posY
            // 
            this.textBox_posY.Location = new System.Drawing.Point(108, 117);
            this.textBox_posY.MinimumSize = new System.Drawing.Size(60, 4);
            this.textBox_posY.Name = "textBox_posY";
            this.textBox_posY.Size = new System.Drawing.Size(67, 20);
            this.textBox_posY.TabIndex = 3;
            this.textBox_posY.Leave += new System.EventHandler(this.TextBox_posY_Leave);
            // 
            // textBox_content
            // 
            this.textBox_content.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_content.Location = new System.Drawing.Point(108, 273);
            this.textBox_content.Name = "textBox_content";
            this.textBox_content.Size = new System.Drawing.Size(516, 20);
            this.textBox_content.TabIndex = 3;
            // 
            // textBox_posX
            // 
            this.textBox_posX.Location = new System.Drawing.Point(108, 78);
            this.textBox_posX.MinimumSize = new System.Drawing.Size(60, 4);
            this.textBox_posX.Name = "textBox_posX";
            this.textBox_posX.Size = new System.Drawing.Size(67, 20);
            this.textBox_posX.TabIndex = 3;
            this.textBox_posX.Leave += new System.EventHandler(this.TextBox_posX_Leave);
            // 
            // comboBox_objectColor
            // 
            this.comboBox_objectColor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_objectColor.FormattingEnabled = true;
            this.comboBox_objectColor.Location = new System.Drawing.Point(181, 115);
            this.comboBox_objectColor.MinimumSize = new System.Drawing.Size(135, 0);
            this.comboBox_objectColor.Name = "comboBox_objectColor";
            this.comboBox_objectColor.Size = new System.Drawing.Size(135, 21);
            this.comboBox_objectColor.TabIndex = 2;
            // 
            // comboBox_fontStyle
            // 
            this.comboBox_fontStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_fontStyle.FormattingEnabled = true;
            this.comboBox_fontStyle.Items.AddRange(new object[] {
            "0 = regular",
            "1 = bold",
            "2 = italic",
            "4 = underline",
            "8 = strikeout"});
            this.comboBox_fontStyle.Location = new System.Drawing.Point(181, 193);
            this.comboBox_fontStyle.MinimumSize = new System.Drawing.Size(135, 0);
            this.comboBox_fontStyle.Name = "comboBox_fontStyle";
            this.comboBox_fontStyle.Size = new System.Drawing.Size(135, 21);
            this.comboBox_fontStyle.TabIndex = 2;
            // 
            // comboBox_object
            // 
            this.comboBox_object.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_object.Enabled = false;
            this.comboBox_object.FormattingEnabled = true;
            this.comboBox_object.Location = new System.Drawing.Point(108, 36);
            this.comboBox_object.MinimumSize = new System.Drawing.Size(60, 0);
            this.comboBox_object.Name = "comboBox_object";
            this.comboBox_object.Size = new System.Drawing.Size(67, 21);
            this.comboBox_object.TabIndex = 2;
            // 
            // comboBox_fontName
            // 
            this.comboBox_fontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_fontName.FormattingEnabled = true;
            this.comboBox_fontName.Location = new System.Drawing.Point(181, 154);
            this.comboBox_fontName.MinimumSize = new System.Drawing.Size(135, 0);
            this.comboBox_fontName.Name = "comboBox_fontName";
            this.comboBox_fontName.Size = new System.Drawing.Size(135, 21);
            this.comboBox_fontName.TabIndex = 2;
            // 
            // comboBox_backgroundColor
            // 
            this.comboBox_backgroundColor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_backgroundColor.FormattingEnabled = true;
            this.comboBox_backgroundColor.Location = new System.Drawing.Point(181, 77);
            this.comboBox_backgroundColor.MinimumSize = new System.Drawing.Size(135, 0);
            this.comboBox_backgroundColor.Name = "comboBox_backgroundColor";
            this.comboBox_backgroundColor.Size = new System.Drawing.Size(135, 21);
            this.comboBox_backgroundColor.TabIndex = 2;
            // 
            // button_clone
            // 
            this.button_clone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_clone.Location = new System.Drawing.Point(3, 240);
            this.button_clone.Name = "button_clone";
            this.button_clone.Size = new System.Drawing.Size(99, 23);
            this.button_clone.TabIndex = 0;
            this.button_clone.Text = "Clone object";
            this.button_clone.UseVisualStyleBackColor = true;
            this.button_clone.Click += new System.EventHandler(this.Button_clone_Click);
            // 
            // button_saveLabel
            // 
            this.button_saveLabel.Location = new System.Drawing.Point(198, 6);
            this.button_saveLabel.Name = "button_saveLabel";
            this.button_saveLabel.Size = new System.Drawing.Size(90, 23);
            this.button_saveLabel.TabIndex = 0;
            this.button_saveLabel.Text = "Save label data";
            this.button_saveLabel.UseVisualStyleBackColor = true;
            this.button_saveLabel.Click += new System.EventHandler(this.Button_saveLabel_Click);
            // 
            // button_saveTemplate
            // 
            this.button_saveTemplate.Location = new System.Drawing.Point(102, 6);
            this.button_saveTemplate.Name = "button_saveTemplate";
            this.button_saveTemplate.Size = new System.Drawing.Size(90, 23);
            this.button_saveTemplate.TabIndex = 0;
            this.button_saveTemplate.Text = "Save template";
            this.button_saveTemplate.UseVisualStyleBackColor = true;
            this.button_saveTemplate.Click += new System.EventHandler(this.Button_saveTemplate_Click);
            // 
            // button_delete
            // 
            this.button_delete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_delete.Location = new System.Drawing.Point(6, 298);
            this.button_delete.Name = "button_delete";
            this.button_delete.Size = new System.Drawing.Size(96, 23);
            this.button_delete.TabIndex = 0;
            this.button_delete.Text = "Delete object";
            this.button_delete.UseVisualStyleBackColor = true;
            this.button_delete.Click += new System.EventHandler(this.Button_delete_Click);
            // 
            // button_down
            // 
            this.button_down.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_down.Location = new System.Drawing.Point(54, 269);
            this.button_down.Name = "button_down";
            this.button_down.Size = new System.Drawing.Size(48, 23);
            this.button_down.TabIndex = 0;
            this.button_down.Text = "Down";
            this.button_down.UseVisualStyleBackColor = true;
            this.button_down.Click += new System.EventHandler(this.Button_down_Click);
            // 
            // button_up
            // 
            this.button_up.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button_up.Location = new System.Drawing.Point(6, 269);
            this.button_up.Name = "button_up";
            this.button_up.Size = new System.Drawing.Size(42, 23);
            this.button_up.TabIndex = 0;
            this.button_up.Text = "Up";
            this.button_up.UseVisualStyleBackColor = true;
            this.button_up.Click += new System.EventHandler(this.Button_up_Click);
            // 
            // button_apply
            // 
            this.button_apply.Location = new System.Drawing.Point(6, 6);
            this.button_apply.Name = "button_apply";
            this.button_apply.Size = new System.Drawing.Size(90, 23);
            this.button_apply.TabIndex = 0;
            this.button_apply.Text = "Apply changes";
            this.button_apply.UseVisualStyleBackColor = true;
            this.button_apply.Click += new System.EventHandler(this.Button_apply_Click);
            // 
            // tabPage_group
            // 
            this.tabPage_group.Controls.Add(this.listBox_objectsMulti);
            this.tabPage_group.Controls.Add(this.button_deleteGroup);
            this.tabPage_group.Controls.Add(this.button_moveLeft);
            this.tabPage_group.Controls.Add(this.button_moveDown);
            this.tabPage_group.Controls.Add(this.button_moveUp);
            this.tabPage_group.Controls.Add(this.textBox_move);
            this.tabPage_group.Controls.Add(this.button_moveRight);
            this.tabPage_group.Location = new System.Drawing.Point(4, 22);
            this.tabPage_group.Name = "tabPage_group";
            this.tabPage_group.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_group.Size = new System.Drawing.Size(633, 334);
            this.tabPage_group.TabIndex = 2;
            this.tabPage_group.Text = "Group operation";
            this.tabPage_group.UseVisualStyleBackColor = true;
            // 
            // listBox_objectsMulti
            // 
            this.listBox_objectsMulti.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.listBox_objectsMulti.FormattingEnabled = true;
            this.listBox_objectsMulti.Location = new System.Drawing.Point(6, 6);
            this.listBox_objectsMulti.Name = "listBox_objectsMulti";
            this.listBox_objectsMulti.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBox_objectsMulti.Size = new System.Drawing.Size(95, 316);
            this.listBox_objectsMulti.TabIndex = 7;
            this.listBox_objectsMulti.SelectedIndexChanged += new System.EventHandler(this.ListBox_objectsMulti_SelectedIndexChanged);
            // 
            // button_deleteGroup
            // 
            this.button_deleteGroup.Location = new System.Drawing.Point(107, 6);
            this.button_deleteGroup.Name = "button_deleteGroup";
            this.button_deleteGroup.Size = new System.Drawing.Size(96, 23);
            this.button_deleteGroup.TabIndex = 5;
            this.button_deleteGroup.Text = "Delete objects";
            this.button_deleteGroup.UseVisualStyleBackColor = true;
            this.button_deleteGroup.Click += new System.EventHandler(this.Button_deleteGroup_Click);
            // 
            // button_moveLeft
            // 
            this.button_moveLeft.Location = new System.Drawing.Point(107, 66);
            this.button_moveLeft.Name = "button_moveLeft";
            this.button_moveLeft.Size = new System.Drawing.Size(41, 23);
            this.button_moveLeft.TabIndex = 0;
            this.button_moveLeft.Text = "Left";
            this.button_moveLeft.UseVisualStyleBackColor = true;
            this.button_moveLeft.Click += new System.EventHandler(this.Button_moveLeft_Click);
            // 
            // button_moveDown
            // 
            this.button_moveDown.Location = new System.Drawing.Point(154, 90);
            this.button_moveDown.Name = "button_moveDown";
            this.button_moveDown.Size = new System.Drawing.Size(50, 23);
            this.button_moveDown.TabIndex = 0;
            this.button_moveDown.Text = "Down";
            this.button_moveDown.UseVisualStyleBackColor = true;
            this.button_moveDown.Click += new System.EventHandler(this.Button_moveDown_Click);
            // 
            // button_moveUp
            // 
            this.button_moveUp.Location = new System.Drawing.Point(154, 35);
            this.button_moveUp.Name = "button_moveUp";
            this.button_moveUp.Size = new System.Drawing.Size(50, 23);
            this.button_moveUp.TabIndex = 0;
            this.button_moveUp.Text = "Up";
            this.button_moveUp.UseVisualStyleBackColor = true;
            this.button_moveUp.Click += new System.EventHandler(this.Button_moveUp_Click);
            // 
            // textBox_move
            // 
            this.textBox_move.Location = new System.Drawing.Point(154, 64);
            this.textBox_move.MinimumSize = new System.Drawing.Size(4, 4);
            this.textBox_move.Name = "textBox_move";
            this.textBox_move.Size = new System.Drawing.Size(50, 20);
            this.textBox_move.TabIndex = 3;
            this.textBox_move.Text = "1";
            this.textBox_move.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_move.Leave += new System.EventHandler(this.TextBox_move_Leave);
            // 
            // button_moveRight
            // 
            this.button_moveRight.Location = new System.Drawing.Point(210, 64);
            this.button_moveRight.Name = "button_moveRight";
            this.button_moveRight.Size = new System.Drawing.Size(41, 23);
            this.button_moveRight.TabIndex = 0;
            this.button_moveRight.Text = "Right";
            this.button_moveRight.UseVisualStyleBackColor = true;
            this.button_moveRight.Click += new System.EventHandler(this.Button_moveRight_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 98);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(57, 13);
            this.label4.TabIndex = 9;
            this.label4.Text = "Resolution";
            // 
            // textBox_dpi
            // 
            this.textBox_dpi.Location = new System.Drawing.Point(60, 96);
            this.textBox_dpi.MinimumSize = new System.Drawing.Size(50, 4);
            this.textBox_dpi.Name = "textBox_dpi";
            this.textBox_dpi.Size = new System.Drawing.Size(50, 20);
            this.textBox_dpi.TabIndex = 8;
            this.textBox_dpi.Text = "200";
            this.textBox_dpi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox_dpi.Leave += new System.EventHandler(this.TextBox_dpi_Leave);
            // 
            // checkBox_columnNames
            // 
            this.checkBox_columnNames.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox_columnNames.AutoSize = true;
            this.checkBox_columnNames.Checked = true;
            this.checkBox_columnNames.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_columnNames.Location = new System.Drawing.Point(486, 43);
            this.checkBox_columnNames.Name = "checkBox_columnNames";
            this.checkBox_columnNames.Size = new System.Drawing.Size(158, 17);
            this.checkBox_columnNames.TabIndex = 2;
            this.checkBox_columnNames.Text = "1st string is a column names";
            this.checkBox_columnNames.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(160, 98);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Units";
            // 
            // textBox_labelsName
            // 
            this.textBox_labelsName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_labelsName.Enabled = false;
            this.textBox_labelsName.Location = new System.Drawing.Point(99, 41);
            this.textBox_labelsName.Name = "textBox_labelsName";
            this.textBox_labelsName.Size = new System.Drawing.Size(381, 20);
            this.textBox_labelsName.TabIndex = 1;
            // 
            // comboBox_units
            // 
            this.comboBox_units.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_units.FormattingEnabled = true;
            this.comboBox_units.Items.AddRange(new object[] {
            "pix",
            "mm",
            "cm",
            "in."});
            this.comboBox_units.Location = new System.Drawing.Point(195, 95);
            this.comboBox_units.MinimumSize = new System.Drawing.Size(50, 0);
            this.comboBox_units.Name = "comboBox_units";
            this.comboBox_units.Size = new System.Drawing.Size(50, 21);
            this.comboBox_units.TabIndex = 6;
            this.comboBox_units.SelectedIndexChanged += new System.EventHandler(this.ComboBox_units_SelectedIndexChanged);
            // 
            // textBox_templateName
            // 
            this.textBox_templateName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_templateName.Enabled = false;
            this.textBox_templateName.Location = new System.Drawing.Point(99, 12);
            this.textBox_templateName.Name = "textBox_templateName";
            this.textBox_templateName.Size = new System.Drawing.Size(545, 20);
            this.textBox_templateName.TabIndex = 1;
            // 
            // button_importTemplate
            // 
            this.button_importTemplate.Location = new System.Drawing.Point(3, 10);
            this.button_importTemplate.Name = "button_importTemplate";
            this.button_importTemplate.Size = new System.Drawing.Size(90, 23);
            this.button_importTemplate.TabIndex = 0;
            this.button_importTemplate.Text = "Import template";
            this.button_importTemplate.UseVisualStyleBackColor = true;
            this.button_importTemplate.Click += new System.EventHandler(this.Button_importTemplate_Click);
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
            this.button_importLabels.Click += new System.EventHandler(this.Button_importLabels_Click);
            // 
            // comboBox_encoding
            // 
            this.comboBox_encoding.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_encoding.FormattingEnabled = true;
            this.comboBox_encoding.Location = new System.Drawing.Point(60, 68);
            this.comboBox_encoding.MinimumSize = new System.Drawing.Size(135, 0);
            this.comboBox_encoding.Name = "comboBox_encoding";
            this.comboBox_encoding.Size = new System.Drawing.Size(185, 21);
            this.comboBox_encoding.TabIndex = 2;
            this.comboBox_encoding.SelectedIndexChanged += new System.EventHandler(this.ComboBox_encoding_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(2, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Encoding";
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
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.Format = "N0";
            dataGridViewCellStyle3.NullValue = null;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_labels.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView_labels.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_labels.Size = new System.Drawing.Size(954, 158);
            this.dataGridView_labels.TabIndex = 0;
            this.dataGridView_labels.SelectionChanged += new System.EventHandler(this.DataGridView_labels_SelectionChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.OpenFileDialog1_FileOk);
            // 
            // printDialog1
            // 
            this.printDialog1.UseEXDialog = true;
            // 
            // SaveFileDialog1
            // 
            this.SaveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.SaveFileDialog1_FileOk);
            // 
            // timer1
            // 
            this.timer1.Interval = 500;
            this.timer1.Tick += new System.EventHandler(this.Timer1_Tick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(958, 654);
            this.Controls.Add(this.splitContainer_horizontal);
            this.MinimumSize = new System.Drawing.Size(620, 600);
            this.Name = "Form1";
            this.Text = "LabelPrint";
            this.splitContainer_horizontal.Panel1.ResumeLayout(false);
            this.splitContainer_horizontal.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer_horizontal)).EndInit();
            this.splitContainer_horizontal.ResumeLayout(false);
            this.splitContainer_vertical.Panel1.ResumeLayout(false);
            this.splitContainer_vertical.Panel1.PerformLayout();
            this.splitContainer_vertical.Panel2.ResumeLayout(false);
            this.splitContainer_vertical.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer_vertical)).EndInit();
            this.splitContainer_vertical.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_label)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage_print.ResumeLayout(false);
            this.tabPage_print.PerformLayout();
            this.tabPage_edit.ResumeLayout(false);
            this.tabPage_edit.PerformLayout();
            this.tabPage_group.ResumeLayout(false);
            this.tabPage_group.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_labels)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer_horizontal;
        private System.Windows.Forms.SplitContainer splitContainer_vertical;
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
        private System.Windows.Forms.CheckBox checkBox_columnNames;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage_print;
        private System.Windows.Forms.TabPage tabPage_edit;
        private System.Windows.Forms.Button button_saveTemplate;
        private System.Windows.Forms.Button button_delete;
        private System.Windows.Forms.Button button_apply;
        private System.Windows.Forms.CheckBox checkBox_fill;
        private System.Windows.Forms.TextBox textBox_height;
        private System.Windows.Forms.TextBox textBox_rotate;
        private System.Windows.Forms.TextBox textBox_width;
        private System.Windows.Forms.TextBox textBox_posY;
        private System.Windows.Forms.TextBox textBox_content;
        private System.Windows.Forms.TextBox textBox_posX;
        private System.Windows.Forms.ComboBox comboBox_objectColor;
        private System.Windows.Forms.ComboBox comboBox_fontStyle;
        private System.Windows.Forms.ComboBox comboBox_fontName;
        private System.Windows.Forms.ComboBox comboBox_backgroundColor;
        private System.Windows.Forms.Label label_content;
        private System.Windows.Forms.Label label_rotate;
        private System.Windows.Forms.Label label_posY;
        private System.Windows.Forms.Label label_height;
        private System.Windows.Forms.Label label_width;
        private System.Windows.Forms.Label label_posX;
        private System.Windows.Forms.Label label_fontName;
        private System.Windows.Forms.Label label_fontStyle;
        private System.Windows.Forms.Label label_objectColor;
        private System.Windows.Forms.Label label_backgroundColor;
        private System.Windows.Forms.Button button_down;
        private System.Windows.Forms.Button button_up;
        private System.Windows.Forms.ListBox listBox_objects;
        private System.Windows.Forms.SaveFileDialog SaveFileDialog1;
        private System.Windows.Forms.ComboBox comboBox_object;
        private System.Windows.Forms.Label label_fontSize;
        private System.Windows.Forms.TextBox textBox_fontSize;
        private System.Windows.Forms.Button button_saveLabel;
        private System.Windows.Forms.ComboBox comboBox_encoding;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox_dpi;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox_units;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button button_clone;
        private System.Windows.Forms.TextBox textBox_mX;
        private System.Windows.Forms.TextBox textBox_mY;
        private System.Windows.Forms.TabPage tabPage_group;
        private System.Windows.Forms.Button button_moveLeft;
        private System.Windows.Forms.Button button_moveDown;
        private System.Windows.Forms.Button button_moveUp;
        private System.Windows.Forms.TextBox textBox_move;
        private System.Windows.Forms.Button button_moveRight;
        private System.Windows.Forms.Button button_deleteGroup;
        private System.Windows.Forms.ListBox listBox_objectsMulti;
    }
}

