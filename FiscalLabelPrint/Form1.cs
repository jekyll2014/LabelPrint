using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;
using ZXing.Common;
using ZXing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Data;
using System.Drawing.Text;
using System.Collections.Generic;

namespace FiscalLabelPrint
{
    public partial class Form1 : Form
    {
        const int labelObject = 0;
        const int textObject = 1;
        const int pictureObject = 2;
        const int barcodeObject = 3;
        const int lineObject = 4;
        const int rectangleObject = 5;
        const int ellipseObject = 6;

        string[] _objectNames = { "label", "text", "picture", "barcode", "line", "rectangle", "ellipse", };
        string[] _textStyleNames = { "0=regular", "1=bold", "2=italic", "4=underline", "8=strikeout" };

        struct template
        {
            public string name;
            public Color bgColor;
            public Color fgColor;
            public float posX;
            public float posY;
            public float rotate;
            public string content;
            public float width;
            public float height;
            public bool transparent;
            public int BCformat;
            public byte fontSize;
            public byte fontStyle;
            public string fontName;
            public string feature;
            public float lineLength;
            public float lineWidth;
            public bool fill;
        }
        List<template> Label = new List<template>();

        DataTable LabelsDatabase = new DataTable();
        int objectsNum = 0;

        List<string> bcFeatures = new List<string> { "AZTEC_LAYERS", "ERROR_CORRECTION", "MARGIN", "PDF417_ASPECT_RATIO", "QR_VERSION" };
        List<int> BarCodeTypes = new List<int> {
        (int)BarcodeFormat.AZTEC,
        (int)BarcodeFormat.CODABAR,
        (int)BarcodeFormat.CODE_128,
        (int)BarcodeFormat.CODE_39,
        (int)BarcodeFormat.CODE_93,
        (int)BarcodeFormat.DATA_MATRIX,
        (int)BarcodeFormat.EAN_13,
        (int)BarcodeFormat.EAN_8,
        (int)BarcodeFormat.IMB,
        (int)BarcodeFormat.ITF,
        (int)BarcodeFormat.MAXICODE,
        (int)BarcodeFormat.MSI,
        (int)BarcodeFormat.PDF_417,
        (int)BarcodeFormat.PLESSEY,
        (int)BarcodeFormat.QR_CODE,
        (int)BarcodeFormat.RSS_14,
        (int)BarcodeFormat.RSS_EXPANDED,
        (int)BarcodeFormat.UPC_A,
        (int)BarcodeFormat.UPC_E,
        (int)BarcodeFormat.UPC_EAN_EXTENSION };

        int pagesFrom = 0;
        int pagesTo = 0;

        bool cmdLinePrint = false;
        string printerName = "";

        Bitmap LabelBmp = new Bitmap(1, 1, PixelFormat.Format32bppPArgb);

        public Form1(string[] cmdLine)
        {
            InitializeComponent();
            if (cmdLine.Length >= 1)
            {
                if (cmdLine[0] == "/?" || cmdLine[0] == "/h" || cmdLine[0] == "/help")
                {
                    Console.WriteLine("/?, /h, /help - print help\r\n" +
                        "/T=file.csv - load template data from file\r\n" +
                        "/L=file.csv - load label data from file\r\n" +
                        "/C - 1st string of label file is column names (default = no)\r\n" +
                        "/PRN=SystemPrinterName - output to printer (replace spaces with \'_\')\r\n" +
                        "/PIC=pictureName - output to pictures\r\n" +
                        "/P=A - print all labels\r\n" +
                        "/P=xxx - print label #xxx\r\n" +
                        "/P=xxx-yyy - print labels from xxx to yyy");
                }
                else
                {
                    string template = "";
                    string label = "";
                    int from = -1;
                    int to = -1;
                    bool printAll = false;

                    for (int i = 0; i < cmdLine.Length; i++)
                    {
                        cmdLine[i] = cmdLine[i].Trim();
                        if (cmdLine[i].ToLower().StartsWith("/t="))
                        {

                            template = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1);
                        }
                        else if (cmdLine[i].ToLower().StartsWith("/c"))
                        {
                            checkBox_columnNames.Checked = true;
                        }
                        else if (cmdLine[i].ToLower().StartsWith("/l="))
                        {
                            label = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1);
                        }
                        else if (cmdLine[i].ToLower().StartsWith("/p="))
                        {
                            cmdLine[i] = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1);
                            if (cmdLine[i] == "a")
                            {
                                printAll = true;
                                from = to = 0;
                            }
                            else if (cmdLine[i].IndexOf('-') > 0)
                            {
                                int.TryParse(cmdLine[i].Substring(0, cmdLine[i].IndexOf('-')), out from);
                                int.TryParse(cmdLine[i].Substring(cmdLine[i].IndexOf('-') + 1), out to);
                            }
                            else
                            {
                                int.TryParse(cmdLine[i], out from);
                                to = from;
                            }
                        }
                        else if (cmdLine[i].ToLower().StartsWith("/prn="))
                        {
                            printerName = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1).Replace("_", " ");
                        }
                        else if (cmdLine[i].ToLower().StartsWith("/pic="))
                        {
                            printerName = " ";
                            textBox_saveFileName.Text = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1);
                            checkBox_toFile.Checked = true;
                        }
                        else
                        {
                            Console.WriteLine("Unknown parameter: " + cmdLine[i]);
                        }
                    }
                    //check we have enough data to print
                    if (template != "" && label != "" && from != -1 && to != -1 && printerName != "")
                    {
                        cmdLinePrint = true;
                        //import template
                        openFileDialog1.Title = "Open template .CSV file";
                        openFileDialog1.FileName = template;
                        openFileDialog1_FileOk(this, null);
                        //import labels
                        openFileDialog1.Title = "Open labels .CSV database";
                        openFileDialog1.FileName = label;
                        openFileDialog1_FileOk(this, null);
                        if (printAll == true) button_printAll_Click(this, EventArgs.Empty);
                        else
                        {
                            textBox_rangeFrom.Text = from.ToString();
                            textBox_rangeTo.Text = to.ToString();
                            button_printRange_Click(this, EventArgs.Empty);
                        }
                    }
                    else Console.WriteLine("Not enough parameters.\r\n");
                }
                if (System.Windows.Forms.Application.MessageLoop)
                {
                    // WinForms app
                    System.Windows.Forms.Application.Exit();
                }
                else
                {
                    // Console app
                    System.Environment.Exit(1);
                }
            }
            comboBox_object.Items.AddRange(_objectNames);
        }

        private void fillBackground(Color bgC)
        {
            LabelBmp = new Bitmap((int)Label[0].width, (int)Label[0].height, PixelFormat.Format32bppPArgb);
            Graphics g = Graphics.FromImage(LabelBmp);
            Rectangle rect = new Rectangle(0, 0, (int)Label[0].width, (int)Label[0].height);
            SolidBrush b = new SolidBrush(bgC);
            g.FillRectangle(b, rect);
            pictureBox_label.Image = LabelBmp;
        }

        private void drawText(Bitmap img, Color fgC, float posX, float posY, string text, string fontName, int fontSize, float rotateDeg = 0, FontStyle fontStyle = FontStyle.Regular)
        {
            Font textFont = new Font(fontName, fontSize, fontStyle); //creates new font
            Graphics g = Graphics.FromImage(pictureBox_label.Image);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            // Save the graphics state.
            GraphicsState state = g.Save();
            g.ResetTransform();
            // Rotate.
            g.RotateTransform(rotateDeg);
            // Translate to desired position. Be sure to append the rotation so it occurs after the rotation.
            g.TranslateTransform(posX, posY, MatrixOrder.Append);
            // Draw the text at the origin.
            SolidBrush b = new SolidBrush(fgC);
            g.DrawString(text, textFont, b, 0, 0);
            // Restore the graphics state.
            g.Restore(state);
        }

        private void drawPicture(Bitmap img, Color fgC, float posX, float posY, string fileName, float rotateDeg = 0, float width = 0, float height = 0, bool makeTransparent = true)
        {
            Bitmap newPicture = new Bitmap(1, 1);
            try
            {
                newPicture = new Bitmap(@fileName, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file: " + fileName + " : " + ex.Message);
                return;
            }
            if (width == 0) width = newPicture.Width;
            if (height == 0) height = newPicture.Height;
            if (makeTransparent) newPicture.MakeTransparent(fgC);
            Graphics g = Graphics.FromImage(img);

            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            GraphicsState state = g.Save();
            g.ResetTransform();
            // Rotate.
            g.RotateTransform(rotateDeg);
            // Translate to desired position. Be sure to append the rotation so it occurs after the rotation.
            g.TranslateTransform(posX, posY, MatrixOrder.Append);
            g.DrawImage(newPicture, 0, 0, width, height);
            // Restore the graphics state.
            g.Restore(state);
        }

        private void drawBarcode(Bitmap img, Color bgC, Color fgC, float posX, float posY, float width, float height, string BCdata, BarcodeFormat bcFormat, float rotateDeg = 0, string addFeature = "", bool makeTransparent = true)
        {
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            GraphicsState state = g.Save();
            g.ResetTransform();
            // Rotate.
            g.RotateTransform(rotateDeg);
            // Translate to desired position. Be sure to append the rotation so it occurs after the rotation.
            g.TranslateTransform(posX, posY, MatrixOrder.Append);
            if (!makeTransparent)
            {
                SolidBrush b = new SolidBrush(bgC);
                Rectangle rect = new Rectangle(0, 0, (int)width, (int)height);
                g.FillRectangle(b, rect);
            }
            var barCode = new BarcodeWriter();
            barCode.Format = bcFormat;
            barCode.Options = new EncodingOptions
            {
                PureBarcode = true,
                Height = (int)height,
                Width = (int)width
            };
            //specify the additional encoding options if there are any
            if (addFeature != "")
            {
                EncodeHintType _feature = EncodeHintType.MIN_SIZE;
                string[] tmp = addFeature.Split('=');
                if (tmp.Length == 2)
                {
                    switch (tmp[0])
                    {
                        case "AZTEC_LAYERS": //int [-4, 32]
                            {
                                _feature = EncodeHintType.AZTEC_LAYERS;
                                int param = 0;
                                int.TryParse(tmp[1], out param);
                                barCode.Options.Hints.Add(_feature, param);
                            }
                            break;
                        case "ERROR_CORRECTION": //int [0, 8]
                            {
                                _feature = EncodeHintType.ERROR_CORRECTION;
                                int param = 0;
                                int.TryParse(tmp[1], out param);
                                barCode.Options.Hints.Add(_feature, param);
                            }
                            break;
                        case "MARGIN":  //int
                            {
                                _feature = EncodeHintType.MARGIN;
                                int param = 0;
                                int.TryParse(tmp[1], out param);
                                barCode.Options.Hints.Add(_feature, param);
                            }
                            break;
                        case "PDF417_ASPECT_RATIO": //int [1, 4]
                            {
                                _feature = EncodeHintType.PDF417_ASPECT_RATIO;
                                int param = 0;
                                int.TryParse(tmp[1], out param);
                                barCode.Options.Hints.Add(_feature, param);
                            }
                            break;
                        case "QR_VERSION": //int [1, 40] ??
                            {
                                _feature = EncodeHintType.QR_VERSION;
                                int param = 0;
                                int.TryParse(tmp[1], out param);
                                barCode.Options.Hints.Add(_feature, param);
                            }
                            break;
                        case "CHARACTER_SET": //string
                            {
                                _feature = EncodeHintType.CHARACTER_SET;
                                barCode.Options.Hints.Add(_feature, tmp[1]);
                            }
                            break;
                        case "PDF417_COMPACTION": //string
                            {
                                _feature = EncodeHintType.PDF417_COMPACTION;
                                barCode.Options.Hints.Add(_feature, tmp[1]);
                            }
                            break;
                        case "CODE128_FORCE_CODESET_B": //bool
                            {
                                _feature = EncodeHintType.CODE128_FORCE_CODESET_B;
                                barCode.Options.Hints.Add(_feature, tmp[1]);
                            }
                            break;
                        case "DISABLE_ECI": //bool
                            {
                                _feature = EncodeHintType.DISABLE_ECI;
                                barCode.Options.Hints.Add(_feature, tmp[1]);
                            }
                            break;
                        case "GS1_FORMAT": //bool
                            {
                                _feature = EncodeHintType.GS1_FORMAT;
                                barCode.Options.Hints.Add(_feature, tmp[1]);
                            }
                            break;
                        case "PDF417_COMPACT": //bool
                            {
                                _feature = EncodeHintType.PDF417_COMPACT;
                                barCode.Options.Hints.Add(_feature, tmp[1]);
                            }
                            break;
                        /*case "DATA_MATRIX_DEFAULT_ENCODATION": //????
                            {
                                _feature = EncodeHintType.DATA_MATRIX_DEFAULT_ENCODATION;
                                barCode.Options.Hints.Add(_feature, tmp[1]);
                            }
                            break;
                        case "DATA_MATRIX_SHAPE": //?????
                            {
                                _feature = EncodeHintType.DATA_MATRIX_SHAPE;
                                barCode.Options.Hints.Add(_feature, tmp[1]);
                            }
                            break;
                        case "PDF417_DIMENSIONS": //????
                            {
                                _feature = EncodeHintType.PDF417_DIMENSIONS;
                                barCode.Options.Hints.Add(_feature, tmp[1]);
                            }
                            break;*/
                        default:
                            MessageBox.Show("Unrecognized additional feature option:" + addFeature);
                            break;
                    }
                }
                else MessageBox.Show("Incorrect additional feature option:" + addFeature);
            }
            Bitmap newPicture = barCode.Write(BCdata);
            newPicture.MakeTransparent(Color.White);
            if (fgC != Color.Black)
            {
                ColorMap[] colorMap = new ColorMap[1];
                colorMap[0] = new ColorMap();
                colorMap[0].OldColor = Color.Black;
                colorMap[0].NewColor = fgC;
                ImageAttributes attr = new ImageAttributes();
                attr.SetRemapTable(colorMap);
                // Draw using the color map
                Rectangle rect = new Rectangle(0, 0, newPicture.Width, newPicture.Height);
                g.DrawImage(newPicture, rect, 0, 0, rect.Width, rect.Height, GraphicsUnit.Pixel, attr);
            }
            else g.DrawImage(newPicture, 0, 0);
            // Restore the graphics state.
            g.Restore(state);
        }

        private void drawLineCoord(Bitmap img, Color fgC, float posX, float posY, float endX, float endY, float lineWidth)
        {
            Pen p = new Pen(fgC, lineWidth);
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawLine(p, posX, posY, endX, endY);
        }

        private void drawLineLength(Bitmap img, Color fgC, float posX, float posY, float length, float rotateDeg, float lineWidth)
        {
            Pen p = new Pen(fgC, lineWidth);
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            // Save the graphics state.
            GraphicsState state = g.Save();
            g.ResetTransform();
            // Rotate.
            g.RotateTransform(rotateDeg);
            // Translate to desired position. Be sure to append the rotation so it occurs after the rotation.
            g.TranslateTransform(posX, posY, MatrixOrder.Append);
            // Draw the line at the origin.
            g.DrawLine(p, 0, 0, length, 0);
            // Restore the graphics state.
            g.Restore(state);
        }

        private void drawRectangle(Bitmap img, Color fgC, float posX, float posY, float width, float height, float rotateDeg, float lineWidth, bool fill)
        {
            Pen p = new Pen(fgC, lineWidth);
            Rectangle rect = new Rectangle(0, 0, (int)width, (int)height);
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            // Save the graphics state.
            GraphicsState state = g.Save();
            g.ResetTransform();
            // Rotate.
            g.RotateTransform(rotateDeg);
            // Translate to desired position. Be sure to append the rotation so it occurs after the rotation.
            g.TranslateTransform(posX, posY, MatrixOrder.Append);
            // Draw the text at the origin.
            g.DrawRectangle(p, rect);
            if (fill)
            {
                SolidBrush b = new SolidBrush(fgC);
                g.FillRectangle(b, rect);
            }
            // Restore the graphics state.
            g.Restore(state);
        }

        private void drawEllipse(Bitmap img, Color fgC, float posX, float posY, float width, float height, float rotateDeg, float lineWidth, bool fill)
        {
            Pen p = new Pen(fgC, lineWidth);
            Rectangle rect = new Rectangle(0, 0, (int)width, (int)height);
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            // Save the graphics state.
            GraphicsState state = g.Save();
            g.ResetTransform();
            // Rotate.
            g.RotateTransform(rotateDeg);
            // Translate to desired position. Be sure to append the rotation so it occurs after the rotation.
            g.TranslateTransform(posX, posY, MatrixOrder.Append);
            // Draw the text at the origin.
            g.DrawEllipse(p, rect);
            if (fill)
            {
                SolidBrush b = new SolidBrush(fgC);
                g.FillEllipse(b, rect);
            }
            // Restore the graphics state.
            g.Restore(state);
        }

        private void ReadCsv(string fileName, DataTable table, bool createColumnsNames = false)
        {
            table.Clear();
            table.Rows.Clear();
            FileStream inputFile;
            try
            {
                inputFile = File.OpenRead(fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file:" + fileName + " : " + ex.Message);
                return;
            }

            //read headers
            StringBuilder inputStr = new StringBuilder();
            int c = 0;
            int colNum = 0;

            if (createColumnsNames == true)
            {
                table.Columns.Clear();
                c = inputFile.ReadByte();
                while (c != '\r' && c != '\n' && c != -1)
                {
                    byte[] b = new byte[1];
                    b[0] = (byte)c;
                    inputStr.Append(Encoding.GetEncoding(LabelPrint.Properties.Settings.Default.CodePage).GetString(b));
                    c = inputFile.ReadByte();
                }

                //create and count columns and read headers
                if (inputStr.Length != 0)
                {
                    string[] cells = inputStr.ToString().Split(LabelPrint.Properties.Settings.Default.CSVdelimiter);
                    colNum = cells.Length - 1;
                    for (int i = 0; i < colNum; i++)
                    {
                        table.Columns.Add(cells[i]);
                    }
                }
            }
            else
            {
                c = inputFile.ReadByte();
                while (c != '\r' && c != '\n' && c != -1)
                {
                    byte[] b = new byte[1];
                    b[0] = (byte)c;
                    inputStr.Append(Encoding.GetEncoding(LabelPrint.Properties.Settings.Default.CodePage).GetString(b));
                    c = inputFile.ReadByte();
                }

                //create 1st row and count columns
                if (inputStr.Length != 0)
                {
                    string[] cells = inputStr.ToString().Split(LabelPrint.Properties.Settings.Default.CSVdelimiter);
                    colNum = cells.Length - 1;
                    DataRow row = table.NewRow();
                    for (int i = 0; i < cells.Length - 1; i++)
                    {
                        string tmp = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').Trim();
                        if (tmp != "") row[i] = tmp;
                    }
                    table.Rows.Add(row);
                }
            }

            //read CSV content string by string
            while (c != -1)
            {
                int i = 0;
                c = 0;
                inputStr.Length = 0;
                while (i < colNum && c != -1 /*&& c != '\r' && c != '\n'*/)
                {
                    c = inputFile.ReadByte();
                    byte[] b = new byte[1];
                    b[0] = (byte)c;
                    if (c == LabelPrint.Properties.Settings.Default.CSVdelimiter) i++;
                    if (c != -1) inputStr.Append(Encoding.GetEncoding(LabelPrint.Properties.Settings.Default.CodePage).GetString(b));
                }
                while (c != '\r' && c != '\n' && c != -1) c = inputFile.ReadByte();
                if (inputStr.ToString().Replace(LabelPrint.Properties.Settings.Default.CSVdelimiter, ' ').Trim().TrimStart('\r').TrimStart('\n').TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').TrimEnd('\r').TrimEnd('\n') != "")
                {
                    string[] cells = inputStr.ToString().Split(LabelPrint.Properties.Settings.Default.CSVdelimiter);

                    DataRow row = table.NewRow();
                    for (i = 0; i < cells.Length - 1; i++)
                    {
                        row[i] = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').Trim();
                    }
                    table.Rows.Add(row);
                }
            }
            inputFile.Close();
        }

        private void button_importLabels_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Title = "Open labels .CSV database";
            openFileDialog1.DefaultExt = "csv";
            openFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
            openFileDialog1.ShowDialog();
        }

        private void button_importTemplate_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Title = "Open template .CSV file";
            openFileDialog1.DefaultExt = "csv";
            openFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (openFileDialog1.Title == "Open labels .CSV database")
            {
                dataGridView_labels.DataSource = null;
                ReadCsv(openFileDialog1.FileName, LabelsDatabase, checkBox_columnNames.Checked);
                if (LabelsDatabase.Rows.Count > 0)
                {
                    dataGridView_labels.DataSource = LabelsDatabase;
                    foreach (DataGridViewColumn column in dataGridView_labels.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
                    //check for picture file existence
                    foreach (DataGridViewRow row in dataGridView_labels.Rows)
                    {
                        for (int i = 0; i < dataGridView_labels.ColumnCount; i++)
                        {
                            if (Label[i + 1].name == _objectNames[pictureObject] && !File.Exists(row.Cells[i].Value.ToString()))
                            {
                                MessageBox.Show("[Line " + (i + 1).ToString() + "] File not exist: " + row.Cells[i].Value.ToString());
                            }
                        }
                    }
                    button_printCurrent.Enabled = true;
                    button_printAll.Enabled = true;
                    button_printRange.Enabled = true;
                    textBox_rangeFrom.Text = "0";
                    textBox_rangeTo.Text = (LabelsDatabase.Rows.Count - 1).ToString();
                    setRowNumber(dataGridView_labels);
                }
                else
                {
                    MessageBox.Show("Error: No label data loaded.");
                    return;
                }
                if (dataGridView_labels.Columns.Count != Label.Count - 1)
                    MessageBox.Show("Label data doesn't match template.\r\nTemplate objects defined:" + (Label.Count - 1).ToString() + "Data loaded: " + dataGridView_labels.Columns.Count.ToString());
                dataGridView_labels.CurrentCell = dataGridView_labels.Rows[0].Cells[0];
                dataGridView_labels.Rows[0].Selected = true;
                generateLabel(dataGridView_labels.CurrentCell.RowIndex);
                textBox_labelsName.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\') + 1);
            }
            else if (openFileDialog1.Title == "Open template .CSV file")
            {
                objectsNum = 0;
                dataGridView_labels.DataSource = null;
                LabelsDatabase.Clear();
                LabelsDatabase.Columns.Clear();
                LabelsDatabase.Rows.Clear();
                textBox_labelsName.Clear();
                string[] inputStr = File.ReadAllLines(openFileDialog1.FileName);
                Label.Clear();
                for (int i = 0; i < inputStr.Length; i++)
                {
                    if (inputStr[i].Trim() != "")
                    {
                        List<string> cells = new List<string>();
                        foreach (string str in inputStr[i].Split(LabelPrint.Properties.Settings.Default.CSVdelimiter))
                        {
                            cells.Add(str.Trim());
                        }
                        if (cells[cells.Count - 1] == "") cells.RemoveAt(cells.Count - 1);
                        if (cells.Count >= 5)
                        {
                            template templ = new template();
                            templ.name = cells[0];
                            objectsNum++;
                            // label; 1 [bgColor]; 2 [objectColor]; 3 width; 4 height;
                            if (templ.name == _objectNames[labelObject])
                            {
                                if (cells[1] != "")
                                {
                                    templ.bgColor = Color.FromName(cells[1]);
                                }
                                else
                                {
                                    templ.bgColor = Color.White;
                                }

                                if (cells[2] != "")
                                {
                                    templ.fgColor = Color.FromName(cells[2]);
                                }
                                else
                                {
                                    templ.fgColor = Color.Black;
                                }
                                if (templ.bgColor == templ.fgColor)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                }

                                float.TryParse(cells[3], out templ.width);
                                if (templ.width <= 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect label width: " + templ.width.ToString());
                                    templ.width = 1;
                                }

                                float.TryParse(cells[4], out templ.height);
                                if (templ.height <= 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect label height: " + templ.height.ToString());
                                    templ.height = 1;
                                }
                            }
                            // text; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [default_text]; 6 fontName; 7 fontSize; 8 [fontStyle];
                            else if (templ.name == _objectNames[textObject])
                            {
                                if (cells.Count >= 9)
                                {
                                    if (cells[1] != "")
                                    {
                                        templ.fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        templ.fgColor = Label[0].fgColor;
                                    }
                                    if (templ.bgColor == templ.fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out templ.posX);
                                    if (templ.posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = (int)Label[0].height - 1;
                                    }

                                    float.TryParse(cells[4], out templ.rotate);

                                    templ.content = cells[5];

                                    templ.fontName = cells[6];
                                    // Check if the font is present
                                    InstalledFontCollection installedFontCollection = new InstalledFontCollection();
                                    bool _fontPresent = false;
                                    foreach (FontFamily fontFamily in installedFontCollection.Families)
                                    {
                                        if (fontFamily.Name == templ.fontName)
                                        {
                                            _fontPresent = true;
                                            break;
                                        }
                                    }
                                    if (_fontPresent == false)
                                    {
                                        MessageBox.Show("Incorrect font name: " + templ.fontName + "\r\nchanged to default: " + this.Font.Name);
                                        templ.fontName = this.Font.Name;
                                    }
                                    installedFontCollection.Dispose();

                                    byte.TryParse(cells[7], out templ.fontSize);

                                    byte.TryParse(cells[8], out templ.fontStyle);
                                    if (templ.fontStyle != 0 && templ.fontStyle != 1 && templ.fontStyle != 2 && templ.fontStyle != 4 && templ.fontStyle != 8)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect font style: " + templ.fontStyle);
                                        templ.fontStyle = 0;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // picture; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [default_file]; 6 [width]; 7 [height]; 8 [transparent];
                            else if (templ.name == _objectNames[pictureObject])
                            {
                                if (cells.Count >= 9)
                                {
                                    if (cells[1] != "")
                                    {
                                        templ.fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        templ.fgColor = Label[0].fgColor;
                                    }
                                    if (templ.bgColor == templ.fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out templ.posX);
                                    if (templ.posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = (int)Label[0].height - 1;
                                    }

                                    float.TryParse(cells[4], out templ.rotate);

                                    templ.content = cells[5];
                                    if (!File.Exists(templ.content))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] File not exist: " + templ.content);
                                    }

                                    float.TryParse(cells[6], out templ.width);
                                    if (templ.width < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = 0;
                                    }
                                    else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }

                                    float.TryParse(cells[7], out templ.height);
                                    if (templ.height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = 0;
                                    }
                                    else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }

                                    byte t = 0;
                                    byte.TryParse(cells[8], out t);
                                    templ.transparent = (t > 0);
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // barcode; 1 [bgColor]; 2 [objectColor]; 3 posX; 4 posY; 5 [rotate]; 6 [default_data]; 7 width; 8 height; 9 bcFormat; 10 [transparent]; 11 [additional_features]
                            else if (templ.name == _objectNames[barcodeObject])
                            {
                                if (cells.Count >= 11)
                                {
                                    if (cells[1] != "")
                                    {
                                        templ.bgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        templ.bgColor = Label[0].bgColor;
                                    }

                                    if (cells[2] != "")
                                    {
                                        templ.fgColor = Color.FromName(cells[2]);
                                    }
                                    else
                                    {
                                        templ.fgColor = Label[0].fgColor;
                                    }
                                    if (templ.bgColor == templ.fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[3], out templ.posX);
                                    if (templ.posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[4], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = (int)Label[0].height - 1;
                                    }

                                    float.TryParse(cells[5], out templ.rotate);

                                    templ.content = cells[6];

                                    float.TryParse(cells[7], out templ.width);
                                    if (templ.width < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = 0;
                                    }
                                    else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }

                                    float.TryParse(cells[8], out templ.height);
                                    if (templ.height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = 0;
                                    }
                                    else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }

                                    int.TryParse(cells[9], out templ.BCformat);
                                    if (!BarCodeTypes.Contains(templ.BCformat))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect barcode type: " + templ.BCformat.ToString());
                                        templ.BCformat = (int)BarcodeFormat.QR_CODE;
                                    }

                                    byte t = 0;
                                    byte.TryParse(cells[10], out t);
                                    templ.transparent = (t > 0);

                                    if (cells.Count >= 12 && cells[11].Contains("="))
                                    {
                                        if (bcFeatures.Contains(cells[11].Substring(0, cells[11].IndexOf('=')).Trim()))
                                        {
                                            templ.feature = cells[11];
                                        }
                                        else if (templ.feature == "0") templ.feature = "";
                                    }
                                    else templ.feature = "";
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // line; 1 [objectColor]; 2 posX; 3 posY; 4 ------- ; 5 [lineWidth]; 6 endX; 7 endY;
                            // line; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 lineLength;
                            else if (templ.name == _objectNames[lineObject])
                            {
                                if (cells.Count >= 7)
                                {
                                    if (cells[1] != "")
                                    {
                                        templ.fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        templ.fgColor = Label[0].fgColor;
                                    }
                                    if (templ.bgColor == templ.fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out templ.posX);
                                    if (templ.posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = (int)Label[0].height - 1;
                                    }

                                    float.TryParse(cells[5], out templ.lineWidth);
                                    if (templ.lineWidth < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect line width: " + templ.lineWidth.ToString());
                                        templ.lineWidth = 0;
                                    }

                                    if (cells[4] != "")
                                    {
                                        float.TryParse(cells[4], out templ.rotate);

                                        float.TryParse(cells[6], out templ.lineLength);
                                        if (templ.lineLength < 0)
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect line length: " + templ.lineLength.ToString());
                                            templ.lineLength = 0;
                                        }
                                        else if (templ.lineLength >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect line length: " + templ.lineLength.ToString());
                                            templ.lineLength = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                        }
                                    }
                                    else
                                    {
                                        float.TryParse(cells[6], out templ.width);
                                        if (templ.width < 0)
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                            templ.width = 0;
                                        }
                                        else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                            templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                        }

                                        float.TryParse(cells[7], out templ.height);
                                        if (templ.height < 0)
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                            templ.height = 0;
                                        }
                                        else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                            templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                        }

                                        templ.lineLength = -1;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // rectangle; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [fill];
                            else if (templ.name == _objectNames[rectangleObject])
                            {
                                if (cells.Count >= 9)
                                {
                                    if (cells[1] != "")
                                    {
                                        templ.fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        templ.fgColor = Label[0].fgColor;
                                    }
                                    if (templ.bgColor == templ.fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out templ.posX);
                                    if (templ.posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = (int)Label[0].height - 1;
                                    }

                                    float.TryParse(cells[4], out templ.rotate);

                                    float.TryParse(cells[5], out templ.lineWidth);
                                    if (templ.lineWidth < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect line width: " + templ.lineWidth.ToString());
                                        templ.lineWidth = 0;
                                    }

                                    float.TryParse(cells[6], out templ.width);
                                    if (templ.width < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = 0;
                                    }
                                    else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }

                                    float.TryParse(cells[7], out templ.height);
                                    if (templ.height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = 0;
                                    }
                                    else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }

                                    byte t = 0;
                                    byte.TryParse(cells[8], out t);
                                    templ.fill = (t > 0);
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // ellipse; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
                            else if (templ.name == _objectNames[ellipseObject])
                            {
                                if (cells.Count >= 9)
                                {
                                    if (cells[1] != "")
                                    {
                                        templ.fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        templ.fgColor = Label[0].fgColor;
                                    }
                                    if (templ.bgColor == templ.fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out templ.posX);
                                    if (templ.posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + templ.posY.ToString());
                                        templ.posY = (int)Label[0].height - 1;
                                    }

                                    float.TryParse(cells[4], out templ.rotate);

                                    float.TryParse(cells[5], out templ.lineWidth);
                                    if (templ.lineWidth < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect line width: " + templ.lineWidth.ToString());
                                        templ.lineWidth = 0;
                                    }

                                    float.TryParse(cells[6], out templ.width);
                                    if (templ.width < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = 0;
                                    }
                                    else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }

                                    float.TryParse(cells[7], out templ.height);
                                    if (templ.height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = 0;
                                    }
                                    else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }

                                    byte t = 0;
                                    byte.TryParse(cells[8], out t);
                                    templ.fill = (t > 0);
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            Label.Add(templ);
                        }
                    }
                }
                button_importLabels.Enabled = true;
                generateLabel(-1);
                textBox_templateName.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\') + 1);
                //create colums and fill 1 row with default values
                for (int i = 1; i < objectsNum; i++)
                {
                    LabelsDatabase.Columns.Add((i).ToString() + " " + Label[i].name);
                }
                DataRow row = LabelsDatabase.NewRow();
                for (int i = 1; i < objectsNum; i++)
                {
                    row[i - 1] = Label[i].content;
                }
                LabelsDatabase.Rows.Add(row);
                dataGridView_labels.DataSource = LabelsDatabase;
                dataGridView_labels.DataSource = LabelsDatabase;
                foreach (DataGridViewColumn column in dataGridView_labels.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
                button_printCurrent.Enabled = true;
                button_printAll.Enabled = false;
                button_printRange.Enabled = false;
                textBox_rangeFrom.Text = "0";
                textBox_rangeTo.Text = "0";
                setRowNumber(dataGridView_labels);
            }
        }

        private void generateLabel(int gridLine)
        {
            if (checkBox_scale.Checked)
            {
                pictureBox_label.Dock = DockStyle.None;
                pictureBox_label.Width = (int)Label[0].width;
                pictureBox_label.Height = (int)Label[0].height;
            }
            else pictureBox_label.Dock = DockStyle.Fill;
            for (int i = 0; i < Label.Count; i++)
            {
                if (Label[i].name == _objectNames[labelObject])
                {
                    LabelBmp = new Bitmap((int)Label[i].width, (int)Label[i].height, PixelFormat.Format32bppPArgb);
                    fillBackground(Label[i].bgColor);
                }
                else if (Label[i].name == _objectNames[textObject])
                {
                    string fontname = Label[i].fontName;
                    int fontSize = Label[i].fontSize;
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    string content = Label[i].content;
                    FontStyle fontStyle = (FontStyle)Label[i].fontStyle;
                    string tmp = "";
                    if (gridLine > -1) tmp = dataGridView_labels.Rows[gridLine].Cells[i - 1].Value.ToString();
                    if (tmp != "") content = tmp;
                    drawText(LabelBmp, Label[i].fgColor, posX, posY, content, fontname, fontSize, rotate, fontStyle);
                }
                else if (Label[i].name == _objectNames[pictureObject])
                {
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    string content = Label[i].content;
                    float width = Label[i].width;
                    float height = Label[i].height;
                    bool transparent = Label[i].transparent;
                    string tmp = "";
                    if (gridLine > -1) tmp = dataGridView_labels.Rows[gridLine].Cells[i - 1].Value.ToString();
                    if (tmp != "") content = tmp;
                    drawPicture(LabelBmp, Label[i].fgColor, posX, posY, content, rotate, width, height, transparent);
                }
                else if (Label[i].name == _objectNames[barcodeObject])
                {
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    string content = Label[i].content;
                    float width = Label[i].width;
                    float height = Label[i].height;
                    bool transparent = Label[i].transparent;
                    BarcodeFormat BCformat = (BarcodeFormat)Label[i].BCformat;
                    string tmp = "";
                    if (gridLine > -1) tmp = dataGridView_labels.Rows[gridLine].Cells[i - 1].Value.ToString();
                    if (tmp != "") content = tmp;
                    string feature = Label[i].feature;

                    drawBarcode(LabelBmp, Label[i].bgColor, Label[i].fgColor, posX, posY, width, height, content, BCformat, rotate, feature, transparent);
                }
                else if (Label[i].name == _objectNames[lineObject])
                {
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    float lineWidth = Label[i].lineWidth;
                    if (Label[i].lineLength == -1)
                    {
                        float endX = Label[i].width;
                        float endY = Label[i].height;
                        drawLineCoord(LabelBmp, Label[i].fgColor, posX, posY, endX, endY, lineWidth);
                    }
                    else
                    {
                        float length = Label[i].lineLength;
                        drawLineLength(LabelBmp, Label[i].fgColor, posX, posY, length, rotate, lineWidth);
                    }
                }
                else if (Label[i].name == _objectNames[rectangleObject])
                {
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    float lineWidth = Label[i].lineWidth;
                    float width = Label[i].width;
                    float height = Label[i].height;
                    bool fill = Label[i].fill;
                    drawRectangle(LabelBmp, Label[i].fgColor, posX, posY, width, height, rotate, lineWidth, fill);
                }
                else if (Label[i].name == _objectNames[ellipseObject])
                {
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    float lineWidth = Label[i].lineWidth;
                    float width = Label[i].width;
                    float height = Label[i].height;
                    bool fill = Label[i].fill;
                    drawEllipse(LabelBmp, Label[i].fgColor, posX, posY, width, height, rotate, lineWidth, fill);
                }
                else MessageBox.Show("Incorrect object: " + Label[i].name);
            }
            pictureBox_label.Image = LabelBmp;
        }

        private void button_printCurrent_Click(object sender, EventArgs e)
        {
            if (pictureBox_label.Image == null) return;
            if (!checkBox_toFile.Checked)
            {
                printDialog1 = new PrintDialog();
                printDocument1 = new PrintDocument();
                printDialog1.Document = printDocument1;
                printDocument1.PrintPage += new PrintPageEventHandler(printImages);
                pagesFrom = dataGridView_labels.CurrentRow.Index;
                pagesTo = pagesFrom;
                if (printDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
            }
            else savePage();
        }

        private void button_printAll_Click(object sender, EventArgs e)
        {
            if (!checkBox_toFile.Checked)
            {
                printDialog1 = new PrintDialog();
                printDocument1 = new PrintDocument();
                printDialog1.Document = printDocument1;
                printDocument1.PrintPage += new PrintPageEventHandler(printImages);
                pagesFrom = 0;
                pagesTo = dataGridView_labels.Rows.Count - 1;
                if (!cmdLinePrint)
                {
                    if (printDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
                }
                else
                {
                    printDocument1.PrinterSettings.PrinterName = printerName;
                    printDocument1.Print();
                }
            }
            else
            {
                for (int i = 0; i < dataGridView_labels.Rows.Count; i++)
                {
                    dataGridView_labels.CurrentCell = dataGridView_labels.Rows[i].Cells[0];
                    dataGridView_labels.Rows[i].Selected = true;
                    generateLabel(dataGridView_labels.CurrentCell.RowIndex);
                    savePage();
                }
            }
        }

        private void button_printRange_Click(object sender, EventArgs e)
        {
            pagesFrom = 0;
            pagesTo = dataGridView_labels.Rows.Count;
            int.TryParse(textBox_rangeFrom.Text, out pagesFrom);
            int.TryParse(textBox_rangeTo.Text, out pagesTo);
            if (pagesFrom < 0) pagesFrom = 0;
            if (pagesTo >= dataGridView_labels.Rows.Count) pagesTo = dataGridView_labels.Rows.Count - 1;

            if (!checkBox_toFile.Checked)
            {
                printDialog1 = new PrintDialog();
                printDocument1 = new PrintDocument();
                printDialog1.Document = printDocument1;
                printDocument1.PrintPage += new PrintPageEventHandler(printImages);
                if (!cmdLinePrint)
                {
                    if (printDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
                }
                else
                {
                    printDocument1.PrinterSettings.PrinterName = printerName;
                    printDocument1.Print();
                }
            }
            else
            {
                for (; pagesFrom <= pagesTo; pagesFrom++)
                {
                    dataGridView_labels.CurrentCell = dataGridView_labels.Rows[pagesFrom].Cells[0];
                    dataGridView_labels.Rows[pagesFrom].Selected = true;
                    generateLabel(dataGridView_labels.CurrentCell.RowIndex);
                    savePage();
                }
            }

        }

        private void printImages(object sender, PrintPageEventArgs args)
        {
            dataGridView_labels.CurrentCell = dataGridView_labels.Rows[pagesFrom].Cells[0];
            dataGridView_labels.Rows[pagesFrom].Selected = true;
            generateLabel(dataGridView_labels.CurrentCell.RowIndex);
            if (pictureBox_label.Image == null) return;
            args.Graphics.PageUnit = GraphicsUnit.Pixel;
            Bitmap bmp = new Bitmap((int)Label[0].width, (int)Label[0].height);
            Rectangle rect = new Rectangle(0, 0, pictureBox_label.Image.Width, pictureBox_label.Image.Height);
            pictureBox_label.DrawToBitmap(bmp, rect);
            args.Graphics.DrawImage(pictureBox_label.Image, rect);
            args.HasMorePages = pagesFrom < pagesTo;
            pagesFrom++;
        }

        private void savePage()
        {
            var dock = pictureBox_label.Dock;
            var sizeMode = pictureBox_label.SizeMode;
            var w = pictureBox_label.Width;
            var h = pictureBox_label.Height;

            pictureBox_label.Dock = DockStyle.None;
            pictureBox_label.SizeMode = PictureBoxSizeMode.CenterImage;
            pictureBox_label.Width = (int)Label[0].width;
            pictureBox_label.Height = (int)Label[0].height;

            Rectangle rect = new Rectangle(0, 0, pictureBox_label.Image.Width, pictureBox_label.Image.Height);
            pictureBox_label.DrawToBitmap(LabelBmp, rect);
            if (LabelBmp != null) LabelBmp.Save(textBox_saveFileName.Text + dataGridView_labels.CurrentCell.RowIndex.ToString() + ".png", ImageFormat.Png);

            pictureBox_label.Dock = dock;
            pictureBox_label.SizeMode = sizeMode;
            pictureBox_label.Width = w;
            pictureBox_label.Height = h;
        }

        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = row.Index.ToString();
            }
            dgv.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void checkBox_toFile_CheckedChanged(object sender, EventArgs e)
        {
            textBox_saveFileName.Enabled = checkBox_toFile.Checked;
        }

        private void checkBox_scale_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_scale.Checked)
            {
                pictureBox_label.Dock = DockStyle.None;
                pictureBox_label.SizeMode = PictureBoxSizeMode.CenterImage;
                pictureBox_label.Width = (int)Label[0].width;
                pictureBox_label.Height = (int)Label[0].height;
            }
            else
            {
                pictureBox_label.SizeMode = PictureBoxSizeMode.Zoom;
                pictureBox_label.Dock = DockStyle.Fill;
            }
        }

        private void dataGridView_labels_SelectionChanged(object sender, EventArgs e)
        {
            if (!cmdLinePrint) generateLabel(dataGridView_labels.CurrentCell.RowIndex);
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
            {
                getObjectsList();
                listBox_objects.SelectedIndex = 0;
            }
        }

        private void listBox_objects_SelectedIndexChanged(object sender, EventArgs e)
        {
            showObject(listBox_objects.SelectedIndex);
        }

        private void button_apply_Click(object sender, EventArgs e)
        {
            int n = listBox_objects.SelectedIndex;
            if (listBox_objects.SelectedIndex == listBox_objects.Items.Count - 1)
            {
                template templ = collectObject();
                Label.Add(templ);
            }
            else
            {
                template templ = collectObject();
                Label[n] = templ;
                generateLabel(-1);
            }
            getObjectsList();
            listBox_objects.SelectedIndex = n;
        }

        private void button_delete_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex != listBox_objects.Items.Count - 1)
            {
                int n = listBox_objects.SelectedIndex;
                Label.RemoveAt(n);
                listBox_objects.Items.Add("");
                getObjectsList();
                listBox_objects.SelectedIndex = n;
            }
        }

        private void button_up_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex != listBox_objects.Items.Count - 1 && listBox_objects.SelectedIndex != 0)
            {
                int n = listBox_objects.SelectedIndex;
                template templ = Label[n];
                Label.RemoveAt(n);
                Label.Insert(n - 1, templ);
                getObjectsList();
                listBox_objects.SelectedIndex = n - 1;
            }
        }

        private void button_down_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < listBox_objects.Items.Count - 2)
            {
                int n = listBox_objects.SelectedIndex;
                template templ = Label[n];
                Label.RemoveAt(n);
                Label.Insert(n + 1, templ);
                getObjectsList();
                listBox_objects.SelectedIndex = n + 1;
            }

        }

        void getObjectsList()
        {
            listBox_objects.Items.Clear();
            foreach (template t in Label)
            {
                listBox_objects.Items.Add(t.name);
            }
            listBox_objects.Items.Add("");
        }

        string[] getColorList()
        {
            List<string> colorList = new List<string>();
            foreach (Color c in new ColorConverter().GetStandardValues())
            {
                colorList.Add(c.Name.ToString());
            }
            return colorList.ToArray();
        }

        string[] getFontList()
        {
            List<string> fontList = new List<string>();
            InstalledFontCollection installedFontCollection = new InstalledFontCollection();
            foreach (FontFamily fontFamily in installedFontCollection.Families)
            {
                fontList.Add(fontFamily.Name);
            }
            installedFontCollection.Dispose();
            return fontList.ToArray();
        }

        string[] getBarcodeList()
        {
            List<string> barcodeList = new List<string>();
            foreach (BarcodeFormat b in BarCodeTypes)
            {
                barcodeList.Add(((int)b).ToString() + "=" + b.ToString());
            }
            return barcodeList.ToArray();
        }

        //****
        void showObject(int n)
        {
            clearFields();
            comboBox_object.SelectedItem = listBox_objects.SelectedItem.ToString();
            string str = "";
            //new object
            if (listBox_objects.SelectedIndex >= Label.Count)
            {
                comboBox_object.Enabled = true;
            }
            // label; [bgColor]; [objectColor]; width; height;
            else if (Label[listBox_objects.SelectedIndex].name == _objectNames[labelObject])
            {
                label_backgroundColor.Text = "Background color";
                comboBox_backgroundColor.Enabled = true;
                comboBox_backgroundColor.Items.Clear();
                if (comboBox_backgroundColor.Enabled)
                {
                    comboBox_backgroundColor.Items.AddRange(getColorList());
                    str = Label[listBox_objects.SelectedIndex].bgColor.Name.ToString();
                    for (int i = 0; i < comboBox_backgroundColor.Items.Count; i++)
                    {
                        if (comboBox_backgroundColor.Items[i].ToString() == str)
                        {
                            comboBox_backgroundColor.SelectedIndex = i;
                            i = comboBox_backgroundColor.Items.Count;
                        }
                    }
                }

                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Enabled = true;
                comboBox_objectColor.Items.Clear();
                if (comboBox_objectColor.Enabled)
                {
                    comboBox_objectColor.Items.Add("");
                    comboBox_objectColor.Items.AddRange(getColorList());
                    str = Label[listBox_objects.SelectedIndex].fgColor.Name.ToString();
                    for (int i = 0; i < comboBox_objectColor.Items.Count; i++)
                    {
                        if (comboBox_objectColor.Items[i].ToString() == str)
                        {
                            comboBox_objectColor.SelectedIndex = i;
                            i = comboBox_objectColor.Items.Count;
                        }
                    }
                }

                label_width.Text = "Width";
                textBox_width.Enabled = true;
                if (textBox_width.Enabled)
                {
                    textBox_width.Text = Label[listBox_objects.SelectedIndex].width.ToString();
                }

                label_height.Text = "Height";
                textBox_height.Enabled = true;
                if (textBox_height.Enabled)
                {
                    textBox_height.Text = Label[listBox_objects.SelectedIndex].height.ToString();
                }
            }

            // text; [objectColor]; posX; posY; [rotate]; [default_text]; fontName; fontSize; [fontStyle];
            else if (Label[listBox_objects.SelectedIndex].name == _objectNames[textObject])
            {
                comboBox_objectColor.Enabled = true;
                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[listBox_objects.SelectedIndex].fgColor.Name.ToString();
                for (int i = 0; i < comboBox_objectColor.Items.Count; i++)
                {
                    if (comboBox_objectColor.Items[i].ToString() == str)
                    {
                        comboBox_objectColor.SelectedIndex = i;
                        i = comboBox_objectColor.Items.Count;
                    }
                }

                textBox_posX.Enabled = true;
                label_posX.Text = "posX";
                textBox_posX.Text = Label[listBox_objects.SelectedIndex].posX.ToString();

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = Label[listBox_objects.SelectedIndex].posY.ToString();

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[listBox_objects.SelectedIndex].rotate.ToString();

                textBox_content.Enabled = true;
                label_content.Text = "Data string";
                textBox_content.Text = Label[listBox_objects.SelectedIndex].content;

                comboBox_fontName.Enabled = true;
                label_fontName.Text = "Font";
                comboBox_fontName.Items.Clear();
                comboBox_fontName.Items.AddRange(getFontList());
                str = Label[listBox_objects.SelectedIndex].fontName.ToString();
                for (int i = 0; i < comboBox_fontName.Items.Count; i++)
                {
                    if (comboBox_fontName.Items[i].ToString() == str)
                    {
                        comboBox_fontName.SelectedIndex = i;
                        i = comboBox_fontName.Items.Count;
                    }
                }

                comboBox_fontStyle.Enabled = true;
                label_fontStyle.Text = "Text style";
                comboBox_fontStyle.Items.Clear();
                comboBox_fontStyle.Items.AddRange(_textStyleNames);
                str = Label[listBox_objects.SelectedIndex].fontStyle.ToString();
                for (int i = 0; i < comboBox_fontStyle.Items.Count; i++)
                {
                    if (comboBox_fontStyle.Items[i].ToString().StartsWith(str))
                    {
                        comboBox_fontStyle.SelectedIndex = i;
                        i = comboBox_fontStyle.Items.Count;
                    }
                }

                textBox_fontSize.Enabled = true;
                label_fontSize.Text = "Font size";
                textBox_fontSize.Text = Label[listBox_objects.SelectedIndex].fontSize.ToString();
            }

            // picture; [objectColor]; posX; posY; [rotate]; [default_file]; [width]; [height]; [transparent];
            else if (Label[listBox_objects.SelectedIndex].name == _objectNames[pictureObject])
            {
                comboBox_objectColor.Enabled = true;
                label_objectColor.Text = "Transparent color";
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[listBox_objects.SelectedIndex].fgColor.Name.ToString();
                for (int i = 0; i < comboBox_objectColor.Items.Count; i++)
                {
                    if (comboBox_objectColor.Items[i].ToString() == str)
                    {
                        comboBox_objectColor.SelectedIndex = i;
                        i = comboBox_objectColor.Items.Count;
                    }
                }

                textBox_posX.Enabled = true;
                label_posX.Text = "posX";
                textBox_posX.Text = Label[listBox_objects.SelectedIndex].posX.ToString();

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = Label[listBox_objects.SelectedIndex].posY.ToString();

                textBox_width.Enabled = true;
                label_width.Text = "Width";
                textBox_width.Text = Label[listBox_objects.SelectedIndex].width.ToString();

                textBox_height.Enabled = true;
                label_height.Text = "Height";
                textBox_height.Text = Label[listBox_objects.SelectedIndex].height.ToString();

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[listBox_objects.SelectedIndex].rotate.ToString();

                textBox_content.Enabled = true;
                label_content.Text = "Picture file";
                textBox_content.Text = Label[listBox_objects.SelectedIndex].content;

                checkBox_fill.Enabled = true;
                checkBox_fill.Text = "Transparent";
                checkBox_fill.Checked = Label[listBox_objects.SelectedIndex].transparent;
            }

            // barcode; [bgColor]; [objectColor]; posX; posY; [rotate]; [default_data]; width; height; bcFormat; [transparent]; [additional_features]
            else if (Label[listBox_objects.SelectedIndex].name == _objectNames[barcodeObject])
            {
                comboBox_backgroundColor.Enabled = true;
                label_backgroundColor.Text = "Background color";
                comboBox_backgroundColor.Items.AddRange(getColorList());
                str = Label[listBox_objects.SelectedIndex].bgColor.Name.ToString();
                for (int i = 0; i < comboBox_backgroundColor.Items.Count; i++)
                {
                    if (comboBox_backgroundColor.Items[i].ToString() == str)
                    {
                        comboBox_backgroundColor.SelectedIndex = i;
                        i = comboBox_backgroundColor.Items.Count;
                    }
                }

                comboBox_objectColor.Enabled = true;
                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[listBox_objects.SelectedIndex].fgColor.Name.ToString();
                for (int i = 0; i < comboBox_objectColor.Items.Count; i++)
                {
                    if (comboBox_objectColor.Items[i].ToString() == str)
                    {
                        comboBox_objectColor.SelectedIndex = i;
                        i = comboBox_objectColor.Items.Count;
                    }
                }

                textBox_posX.Enabled = true;
                label_posX.Text = "posX";
                textBox_posX.Text = Label[listBox_objects.SelectedIndex].posX.ToString();

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = Label[listBox_objects.SelectedIndex].posY.ToString();

                textBox_width.Enabled = true;
                label_width.Text = "Width";
                textBox_width.Text = Label[listBox_objects.SelectedIndex].width.ToString();

                textBox_height.Enabled = true;
                label_height.Text = "Height";
                textBox_height.Text = Label[listBox_objects.SelectedIndex].height.ToString();

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[listBox_objects.SelectedIndex].rotate.ToString();

                textBox_content.Enabled = true;
                label_content.Text = "Data string";
                textBox_content.Text = Label[listBox_objects.SelectedIndex].content;

                comboBox_fontName.Enabled = true;
                label_fontName.Text = "Barcode type";
                comboBox_fontName.Items.Clear();
                comboBox_fontName.Items.AddRange(getBarcodeList());
                str = Label[listBox_objects.SelectedIndex].BCformat.ToString() + "=" + ((BarcodeFormat)Label[listBox_objects.SelectedIndex].BCformat).ToString();
                for (int i = 0; i < comboBox_fontName.Items.Count; i++)
                {
                    if (comboBox_fontName.Items[i].ToString() == str)
                    {
                        comboBox_fontName.SelectedIndex = i;
                        i = comboBox_fontName.Items.Count;
                    }
                }

                comboBox_fontStyle.Enabled = true;
                label_fontStyle.Text = "Additional feature";
                comboBox_fontStyle.Items.AddRange(bcFeatures.ToArray());
                str = Label[listBox_objects.SelectedIndex].feature;
                for (int i = 0; i < comboBox_fontStyle.Items.Count; i++)
                {
                    if (str.StartsWith(comboBox_fontStyle.Items[i].ToString()))
                    {
                        comboBox_fontStyle.SelectedIndex = i;
                    }
                }

                textBox_fontSize.Enabled = true;
                label_fontSize.Text = "Feature value";
                textBox_fontSize.Text = Label[listBox_objects.SelectedIndex].feature.Substring(Label[listBox_objects.SelectedIndex].feature.IndexOf('=') + 1);

                checkBox_fill.Enabled = true;
                checkBox_fill.Text = "Transparent";
                checkBox_fill.Checked = Label[listBox_objects.SelectedIndex].transparent;
            }

            // line; [objectColor]; posX; posY; --- ; [lineWidth]; endX; endY; (lineLength = -1)
            // line; [objectColor]; posX; posY; [rotate]; [lineWidth]; lineLength;
            else if (Label[listBox_objects.SelectedIndex].name == _objectNames[lineObject])
            {
                comboBox_objectColor.Enabled = true;
                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[listBox_objects.SelectedIndex].fgColor.Name.ToString();
                for (int i = 0; i < comboBox_objectColor.Items.Count; i++)
                {
                    if (comboBox_objectColor.Items[i].ToString() == str)
                    {
                        comboBox_objectColor.SelectedIndex = i;
                        i = comboBox_objectColor.Items.Count;
                    }
                }

                textBox_posX.Enabled = true;
                label_posX.Text = "posX";
                textBox_posX.Text = Label[listBox_objects.SelectedIndex].posX.ToString();

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = Label[listBox_objects.SelectedIndex].posY.ToString();

                textBox_fontSize.Enabled = true;
                label_fontSize.Text = "Line width";
                textBox_fontSize.Text = Label[listBox_objects.SelectedIndex].lineWidth.ToString();

                if (Label[listBox_objects.SelectedIndex].lineLength == -1)
                {
                    textBox_width.Enabled = true;
                    label_width.Text = "endX";
                    textBox_width.Text = Label[listBox_objects.SelectedIndex].width.ToString();

                    label_height.Text = "endY";
                    textBox_height.Enabled = true;
                    textBox_height.Text = Label[listBox_objects.SelectedIndex].height.ToString();
                }
                else
                {
                    textBox_rotate.Enabled = true;
                    label_rotate.Text = "Rotate";
                    textBox_rotate.Text = Label[listBox_objects.SelectedIndex].rotate.ToString();

                    textBox_content.Enabled = true;
                    textBox_content.Text = "Line length";
                    textBox_content.Text = Label[listBox_objects.SelectedIndex].lineLength.ToString();
                }
            }

            // rectangle; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
            else if (Label[listBox_objects.SelectedIndex].name == _objectNames[rectangleObject])
            {
                comboBox_objectColor.Enabled = true;
                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[listBox_objects.SelectedIndex].fgColor.Name.ToString();
                for (int i = 0; i < comboBox_objectColor.Items.Count; i++)
                {
                    if (comboBox_objectColor.Items[i].ToString() == str)
                    {
                        comboBox_objectColor.SelectedIndex = i;
                        i = comboBox_objectColor.Items.Count;
                    }
                }

                textBox_posX.Enabled = true;
                label_posX.Text = "posX";
                textBox_posX.Text = Label[listBox_objects.SelectedIndex].posX.ToString();

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = Label[listBox_objects.SelectedIndex].posY.ToString();

                textBox_width.Enabled = true;
                label_width.Text = "Width";
                textBox_width.Text = Label[listBox_objects.SelectedIndex].width.ToString();

                textBox_height.Enabled = true;
                label_height.Text = "Height";
                textBox_height.Text = Label[listBox_objects.SelectedIndex].height.ToString();

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[listBox_objects.SelectedIndex].rotate.ToString();

                textBox_fontSize.Enabled = true;
                label_fontSize.Text = "Line width";
                textBox_fontSize.Text = Label[listBox_objects.SelectedIndex].lineWidth.ToString();

                textBox_content.Enabled = true;
                textBox_content.Text = "Line length";
                textBox_content.Text = Label[listBox_objects.SelectedIndex].lineLength.ToString();

                checkBox_fill.Enabled = true;
                checkBox_fill.Text = "Fill with objectColor";
                checkBox_fill.Checked = Label[listBox_objects.SelectedIndex].fill;
            }

            // ellipse; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
            else if (Label[listBox_objects.SelectedIndex].name == _objectNames[ellipseObject])
            {
                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Enabled = true;
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[listBox_objects.SelectedIndex].fgColor.Name.ToString();
                for (int i = 0; i < comboBox_objectColor.Items.Count; i++)
                {
                    if (comboBox_objectColor.Items[i].ToString() == str)
                    {
                        comboBox_objectColor.SelectedIndex = i;
                        i = comboBox_objectColor.Items.Count;
                    }
                }

                textBox_posX.Enabled = true;
                label_posX.Text = "posX";
                textBox_posX.Text = Label[listBox_objects.SelectedIndex].posX.ToString();

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = Label[listBox_objects.SelectedIndex].posY.ToString();

                textBox_width.Enabled = true;
                label_width.Text = "Width";
                textBox_width.Text = Label[listBox_objects.SelectedIndex].width.ToString();

                textBox_height.Enabled = true;
                label_height.Text = "Height";
                textBox_height.Text = Label[listBox_objects.SelectedIndex].height.ToString();

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[listBox_objects.SelectedIndex].rotate.ToString();

                textBox_fontSize.Enabled = true;
                label_fontSize.Text = "Line width";
                textBox_fontSize.Text = Label[listBox_objects.SelectedIndex].lineWidth.ToString();

                textBox_content.Enabled = true;
                textBox_content.Text = "Line length";
                textBox_content.Text = Label[listBox_objects.SelectedIndex].lineLength.ToString();

                checkBox_fill.Enabled = true;
                checkBox_fill.Text = "Fill with objectColor";
                checkBox_fill.Checked = Label[listBox_objects.SelectedIndex].fill;
            }
        }

        //****
        template collectObject()
        {
            template templ = new template();

            templ.name = comboBox_object.SelectedItem.ToString();

            if (listBox_objects.SelectedIndex == listBox_objects.Items.Count - 1)
            {
                templ.bgColor = Color.White;
                templ.fgColor = Color.Black;
                templ.posX = 0;
                templ.posY = 0;
                templ.rotate = 0;
                templ.content = "";
                templ.width = 1;
                templ.height = 1;
                templ.transparent = false;
                templ.BCformat = 1;
                templ.fontSize = 1;
                templ.fontStyle = 0;
                templ.fontName = "";
                templ.feature = "";
                templ.lineLength = 1;
                templ.lineWidth = 1;
                templ.fill = false;
            }
            // label; [bgColor]; [objectColor]; width; height;
            else if (templ.name == _objectNames[labelObject])
            {
                if (comboBox_backgroundColor.SelectedItem.ToString() == "Background color") templ.bgColor = Label[0].bgColor;
                else templ.bgColor = Color.FromName(comboBox_backgroundColor.SelectedItem.ToString());

                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_width.Text, out f);
                templ.width = f;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f;
            }

            // text; [objectColor]; posX; posY; [rotate]; [default_text]; fontName; fontSize; [fontStyle];
            else if (templ.name == _objectNames[textObject])
            {
                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.bgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                templ.content = textBox_content.Text;

                byte b = 0;
                byte.TryParse(textBox_fontSize.Text, out b);
                templ.fontSize = b;

                byte.TryParse(comboBox_fontStyle.SelectedItem.ToString().Substring(0, 1), out b);
                templ.fontStyle = b;

                templ.fontName = comboBox_fontName.SelectedItem.ToString();
            }

            // picture; [objectColor]; posX; posY; [rotate]; [default_file]; [width]; [height]; [transparent];
            else if (templ.name == _objectNames[pictureObject])
            {
                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.bgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                templ.content = textBox_content.Text;

                float.TryParse(textBox_width.Text, out f);
                templ.width = f;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f;

                templ.transparent = checkBox_fill.Checked;
            }

            // barcode; [bgColor]; [objectColor]; posX; posY; [rotate]; [default_data]; width; height; bcFormat; [transparent]; [additional_features]
            else if (templ.name == _objectNames[barcodeObject])
            {
                if (comboBox_backgroundColor.SelectedItem.ToString() == "Background color") templ.bgColor = Label[0].bgColor;
                else templ.bgColor = Color.FromName(comboBox_backgroundColor.SelectedItem.ToString());

                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                templ.content = textBox_content.Text;

                float.TryParse(textBox_width.Text, out f);
                templ.width = f;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f;

                templ.transparent = checkBox_fill.Checked;

                int i = 0;
                int.TryParse(comboBox_fontName.SelectedItem.ToString().Substring(0, comboBox_fontName.SelectedItem.ToString().IndexOf('=')), out i);
                templ.BCformat = i;

                if (comboBox_fontStyle.SelectedItem.ToString() != "") templ.feature = comboBox_fontStyle.SelectedItem.ToString() + "=" + textBox_fontSize.Text;
                else templ.feature = "";
            }

            // line; [objectColor]; posX; posY; --- ; [lineWidth]; endX; endY; (lineLength = -1)
            // line; [objectColor]; posX; posY; [rotate]; [lineWidth]; lineLength;
            else if (templ.name == _objectNames[lineObject])
            {
                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f;

                float.TryParse(textBox_fontSize.Text, out f);
                templ.lineWidth = f;

                if (textBox_rotate.Text == "")
                {
                    float.TryParse(textBox_width.Text, out f);
                    templ.width = f;

                    float.TryParse(textBox_height.Text, out f);
                    templ.height = f;

                    templ.lineLength = -1;
                }
                else
                {
                    float.TryParse(textBox_rotate.Text, out f);
                    templ.rotate = f;

                    float.TryParse(textBox_content.Text, out f);
                    templ.lineLength = f;
                }
            }

            // rectangle; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
            else if (templ.name == _objectNames[rectangleObject])
            {
                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                float.TryParse(textBox_width.Text, out f);
                templ.width = f;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f;

                float.TryParse(textBox_fontSize.Text, out f);
                templ.lineWidth = f;

                templ.fill = checkBox_fill.Checked;
            }

            // ellipse; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
            else if (templ.name == _objectNames[ellipseObject])
            {
                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                float.TryParse(textBox_width.Text, out f);
                templ.width = f;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f;

                float.TryParse(textBox_fontSize.Text, out f);
                templ.lineWidth = f;

                templ.fill = checkBox_fill.Checked;
            }

            return templ;
        }

        void clearFields()
        {
            comboBox_object.Enabled = false;

            label_backgroundColor.Text = "Background color";
            comboBox_backgroundColor.Items.Clear();
            comboBox_backgroundColor.Items.Add("Default background");
            comboBox_backgroundColor.SelectedIndex = 0;
            comboBox_backgroundColor.Enabled = false;

            label_objectColor.Text = "Default object color";
            comboBox_objectColor.Items.Clear();
            comboBox_objectColor.Items.Add("Default objectColor");
            comboBox_objectColor.SelectedIndex = 0; ;
            comboBox_objectColor.Enabled = false;

            label_posX.Text = "posX";
            textBox_posX.Clear();
            textBox_posX.Enabled = false;

            label_posY.Text = "posY";
            textBox_posY.Clear();
            textBox_posY.Enabled = false;

            label_width.Text = "Width";
            textBox_width.Clear();
            textBox_width.Enabled = false;

            label_height.Text = "Height";
            textBox_height.Clear();
            textBox_height.Enabled = false;

            label_rotate.Text = "Rotate";
            textBox_rotate.Clear();
            textBox_rotate.Enabled = false;

            label_content.Text = "Data string";
            textBox_content.Clear();
            textBox_content.Enabled = false;

            label_fontName.Text = "Font";
            comboBox_fontName.Items.Clear();
            comboBox_fontName.Items.Add("");
            comboBox_fontName.SelectedIndex = 0;
            comboBox_fontName.Enabled = false;

            label_fontStyle.Text = "Text style";
            comboBox_fontStyle.Items.Clear();
            comboBox_fontStyle.Items.Add("");
            comboBox_fontStyle.SelectedIndex = 0;
            comboBox_fontStyle.Enabled = false;

            label_fontSize.Text = "Font size";
            textBox_fontSize.Clear();
            textBox_fontSize.Enabled = false;

            checkBox_fill.Enabled = false;
            checkBox_fill.Text = "Transparent";

        }

        private void button_save_Click(object sender, EventArgs e)
        {
            SaveFileDialog1.Title = "Save .CSV table as...";
            SaveFileDialog1.DefaultExt = "csv";
            SaveFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
            SaveFileDialog1.FileName = "production_" + DateTime.Today.ToShortDateString().Replace("/", "_") + ".csv";
            SaveFileDialog1.ShowDialog();
        }

        private void saveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StringBuilder output = new StringBuilder();
            for (int i = 0; i < Label.Count; i++)
            {
                char div = LabelPrint.Properties.Settings.Default.CSVdelimiter;
                // label; 1 [bgColor]; 2 [objectColor]; 3 width; 4 height;
                if (Label[listBox_objects.SelectedIndex].name == _objectNames[labelObject])
                {
                    output.AppendLine(Label[i].bgColor.Name.ToString() + div +
                        Label[i].fgColor.Name.ToString() + div +
                        Label[i].width.ToString() + div +
                        Label[i].height.ToString() + div);
                }
                // text; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [default_text]; 6 fontName; 7 fontSize; 8 [fontStyle];
                else if (Label[listBox_objects.SelectedIndex].name == _objectNames[textObject])
                {
                    output.AppendLine(Label[i].fgColor.Name.ToString() + div +
                        Label[i].posX.ToString() + div +
                        Label[i].posY.ToString() + div +
                        Label[i].rotate.ToString() + div +
                        Label[i].content.ToString() + div +
                        Label[i].fontName.ToString() + div +
                        Label[i].fontSize.ToString() + div +
                        Label[i].fontStyle.ToString() + div);
                }
                // picture; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [default_file]; 6 [width]; 7 [height]; 8 [transparent];
                else if (Label[listBox_objects.SelectedIndex].name == _objectNames[pictureObject])
                {
                    output.AppendLine(Label[i].fgColor.Name.ToString() + div +
                        Label[i].posX.ToString() + div +
                        Label[i].posY.ToString() + div +
                        Label[i].rotate.ToString() + div +
                        Label[i].content.ToString() + div +
                        Label[i].width.ToString() + div +
                        Label[i].height.ToString() + div +
                        Label[i].transparent.ToString() + div);
                }
                // barcode; 1 [bgColor]; 2 [objectColor]; 3 posX; 4 posY; 5 [rotate]; 6 [default_data]; 7 width; 8 height; 9 bcFormat; 10 [transparent]; 11 [additional_features]
                else if (Label[listBox_objects.SelectedIndex].name == _objectNames[barcodeObject])
                {
                    output.AppendLine(Label[i].bgColor.Name.ToString() + div +
                        Label[i].fgColor.Name.ToString() + div +
                        Label[i].posX.ToString() + div +
                        Label[i].posY.ToString() + div +
                        Label[i].rotate.ToString() + div +
                        Label[i].content.ToString() + div +
                        Label[i].width.ToString() + div +
                        Label[i].height.ToString() + div +
                        Label[i].BCformat.ToString() + div +
                        Label[i].transparent.ToString() + div +
                        Label[i].feature.ToString() + div);
                }
                // line; 1 [objectColor]; 2 posX; 3 posY; 4 ------- ; 5 [lineWidth]; 6 endX; 7 endY; (lineLength = -1)
                // line; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 lineLength;
                else if (Label[listBox_objects.SelectedIndex].name == _objectNames[lineObject])
                {
                    output.Append(
                        Label[i].fgColor.Name.ToString() + div +
                            Label[i].posX.ToString() + div +
                            Label[i].posY.ToString() + div);
                    if (Label[i].lineLength == -1)
                    {
                        output.AppendLine("" + div +
                            Label[i].lineWidth.ToString() + div +
                            Label[i].width.ToString() + div +
                            Label[i].height.ToString() + div);
                    }
                    else
                    {
                        output.AppendLine(Label[i].rotate.ToString() + div +
                            Label[i].lineWidth.ToString() + div +
                            Label[i].lineLength.ToString() + div);
                    }
                }
                // rectangle; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [fill];
                else if (Label[listBox_objects.SelectedIndex].name == _objectNames[rectangleObject])
                {
                    output.AppendLine(Label[i].fgColor.Name.ToString() + div +
                        Label[i].posX.ToString() + div +
                        Label[i].posY.ToString() + div +
                        Label[i].rotate.ToString() + div +
                        Label[i].lineWidth.ToString() + div +
                        Label[i].width.ToString() + div +
                        Label[i].height.ToString() + div +
                        Label[i].fill.ToString() + div);
                }
                // ellipse; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [fill];
                else if (Label[listBox_objects.SelectedIndex].name == _objectNames[ellipseObject])
                {
                    output.AppendLine(Label[i].fgColor.Name.ToString() + div +
                        Label[i].posX.ToString() + div +
                        Label[i].posY.ToString() + div +
                        Label[i].rotate.ToString() + div +
                        Label[i].lineWidth.ToString() + div +
                        Label[i].width.ToString() + div +
                        Label[i].height.ToString() + div +
                        Label[i].fill.ToString() + div);
                }
            }
            try
            {
                File.WriteAllText(SaveFileDialog1.FileName, output.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error writing to file " + SaveFileDialog1.FileName + ": " + ex.Message);
            }
        }
    }
}
