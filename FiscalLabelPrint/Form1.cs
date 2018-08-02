﻿using System;
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

        struct template
        {
            public string type;
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
        template[] Label = new template[0];

        DataTable LabelsDatabase = new DataTable();
        int objectsNum = 0;

        int[] BarCodeTypes = {
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

        int labelWidth = 0;
        int labelHeight = 0;
        Color bgColor = Color.White;
        Color fgColor = Color.Black;

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

        }

        private void fillBackground(Color bgC)
        {
            LabelBmp = new Bitmap(labelWidth, labelHeight, PixelFormat.Format32bppPArgb);
            Graphics g = Graphics.FromImage(LabelBmp);
            Rectangle rect = new Rectangle(0, 0, labelWidth, labelHeight);
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
                            if (Label[i + 1].type == _objectNames[pictureObject] && !File.Exists(row.Cells[i].Value.ToString()))
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
                if (dataGridView_labels.Columns.Count != Label.Length - 1)
                    MessageBox.Show("Label data doesn't match template.\r\nTemplate objects defined:" + (Label.Length - 1).ToString() + "Data loaded: " + dataGridView_labels.Columns.Count.ToString());
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
                Label = new template[inputStr.Length];
                for (int i = 0; i < inputStr.Length; i++)
                {
                    if (inputStr[i].Trim() != "")
                    {
                        string[] cells = inputStr[i].Split(LabelPrint.Properties.Settings.Default.CSVdelimiter);
                        for (int i1 = 0; i1 < cells.Length; i1++) cells[i1] = cells[i1].Trim();
                        if (cells.Length >= 5)
                        {
                            Label[i].type = cells[0];
                            objectsNum++;
                            // label; [bgColor]; [objectColor]; width; height;
                            if (Label[i].type == _objectNames[labelObject])
                            {
                                if (cells[1] != "")
                                {
                                    Label[i].bgColor = bgColor = Color.FromName(cells[1]);
                                }
                                else
                                {
                                    Label[i].bgColor = bgColor = Color.White;
                                }

                                if (cells[2] != "")
                                {
                                    Label[i].fgColor = fgColor = Color.FromName(cells[2]);
                                }
                                else
                                {
                                    Label[i].fgColor = fgColor = Color.Black;
                                }
                                if (Label[i].bgColor == Label[i].fgColor)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                }

                                float.TryParse(cells[3], out Label[i].width);
                                if (Label[i].width <= 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect label width: " + Label[i].width.ToString());
                                    Label[i].width = 1;
                                }
                                labelWidth = (int)Label[i].width;

                                float.TryParse(cells[4], out Label[i].height);
                                if (Label[i].height <= 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect label height: " + Label[i].height.ToString());
                                    Label[i].height = 1;
                                }
                                labelHeight = (int)Label[i].height;
                            }
                            // text; [objectColor]; posX; posY; [rotate]; [default_text]; fontName; fontSize; [fontStyle];
                            else if (Label[i].type == _objectNames[textObject])
                            {
                                if (cells.Length >= 9)
                                {
                                    if (cells[1] != "")
                                    {
                                        Label[i].fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        Label[i].fgColor = fgColor;
                                    }
                                    if (Label[i].bgColor == Label[i].fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out Label[i].posX);
                                    if (Label[i].posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = 0;
                                    }
                                    else if (Label[i].posX >= labelWidth)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = labelWidth - 1;
                                    }

                                    float.TryParse(cells[3], out Label[i].posY);
                                    if (Label[i].posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = 0;
                                    }
                                    else if (Label[i].posY >= labelHeight)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = labelHeight - 1;
                                    }

                                    float.TryParse(cells[4], out Label[i].rotate);

                                    Label[i].content = cells[5];

                                    Label[i].fontName = cells[6];
                                    // Check if the font is present
                                    InstalledFontCollection installedFontCollection = new InstalledFontCollection();
                                    bool _fontPresent = false;
                                    foreach (FontFamily fontFamily in installedFontCollection.Families)
                                    {
                                        if (fontFamily.Name == Label[i].fontName)
                                        {
                                            _fontPresent = true;
                                            break;
                                        }
                                    }
                                    if (_fontPresent == false)
                                    {
                                        MessageBox.Show("Incorrect font name: " + Label[i].fontName + "\r\nchanged to default: " + this.Font.Name);
                                        Label[i].fontName = this.Font.Name;
                                    }
                                    installedFontCollection.Dispose();

                                    byte.TryParse(cells[7], out Label[i].fontSize);

                                    byte.TryParse(cells[8], out Label[i].fontStyle);
                                    if (Label[i].fontStyle != 0 && Label[i].fontStyle != 1 && Label[i].fontStyle != 2 && Label[i].fontStyle != 4 && Label[i].fontStyle != 8)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect font style: " + Label[i].fontStyle);
                                        Label[i].fontStyle = 0;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // picture; [objectColor]; posX; posY; [rotate]; [default_file]; [width]; [height]; [transparent];
                            else if (Label[i].type == _objectNames[pictureObject])
                            {
                                if (cells.Length >= 9)
                                {
                                    if (cells[1] != "")
                                    {
                                        Label[i].fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        Label[i].fgColor = fgColor;
                                    }
                                    if (Label[i].bgColor == Label[i].fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out Label[i].posX);
                                    if (Label[i].posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = 0;
                                    }
                                    else if (Label[i].posX >= labelWidth)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = labelWidth - 1;
                                    }

                                    float.TryParse(cells[3], out Label[i].posY);
                                    if (Label[i].posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = 0;
                                    }
                                    else if (Label[i].posY >= labelHeight)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = labelHeight - 1;
                                    }

                                    float.TryParse(cells[4], out Label[i].rotate);

                                    Label[i].content = cells[5];
                                    if (!File.Exists(Label[i].content))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] File not exist: " + Label[i].content);
                                    }

                                    float.TryParse(cells[6], out Label[i].width);
                                    if (Label[i].width < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                        Label[i].width = 0;
                                    }
                                    else if (Label[i].width >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                        Label[i].width = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                    }

                                    float.TryParse(cells[7], out Label[i].height);
                                    if (Label[i].height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                        Label[i].height = 0;
                                    }
                                    else if (Label[i].height >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                        Label[i].height = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                    }

                                    byte t = 0;
                                    byte.TryParse(cells[8], out t);
                                    Label[i].transparent = (t > 0);
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // barcode; [bgColor]; [objectColor]; posX; posY; [rotate]; [default_data]; width; height; bcFormat; [transparent]; [additional_features]
                            else if (Label[i].type == _objectNames[barcodeObject])
                            {
                                if (cells.Length >= 11)
                                {
                                    if (cells[1] != "")
                                    {
                                        Label[i].bgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        Label[i].bgColor = bgColor;
                                    }

                                    if (cells[2] != "")
                                    {
                                        Label[i].fgColor = Color.FromName(cells[2]);
                                    }
                                    else
                                    {
                                        Label[i].fgColor = fgColor;
                                    }
                                    if (Label[i].bgColor == Label[i].fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[3], out Label[i].posX);
                                    if (Label[i].posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = 0;
                                    }
                                    else if (Label[i].posX >= labelWidth)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = labelWidth - 1;
                                    }

                                    float.TryParse(cells[4], out Label[i].posY);
                                    if (Label[i].posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = 0;
                                    }
                                    else if (Label[i].posY >= labelHeight)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = labelHeight - 1;
                                    }

                                    float.TryParse(cells[5], out Label[i].rotate);

                                    Label[i].content = cells[6];

                                    float.TryParse(cells[7], out Label[i].width);
                                    if (Label[i].width < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                        Label[i].width = 0;
                                    }
                                    else if (Label[i].width >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                        Label[i].width = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                    }

                                    float.TryParse(cells[8], out Label[i].height);
                                    if (Label[i].height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                        Label[i].height = 0;
                                    }
                                    else if (Label[i].height >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                        Label[i].height = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                    }

                                    int.TryParse(cells[9], out Label[i].BCformat);
                                    Boolean _bcCorrect = false;
                                    for (int i1 = 0; i1 < BarCodeTypes.Length; i1++)
                                    {
                                        if (BarCodeTypes[i1] == Label[i].BCformat)
                                        {
                                            _bcCorrect = true;
                                            break;
                                        }
                                    }
                                    if (!_bcCorrect)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect barcode type: " + Label[i].BCformat.ToString());
                                        Label[i].BCformat = (int)BarcodeFormat.QR_CODE;
                                    }

                                    byte t = 0;
                                    byte.TryParse(cells[10], out t);
                                    Label[i].transparent = (t > 0);

                                    if (cells.Length >= 12)
                                    {
                                        Label[i].feature = cells[11];
                                        if (Label[i].feature == "0") Label[i].feature = "";
                                    }
                                    else Label[i].feature = "";
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // line; [objectColor]; posX; posY; --- ; [lineWidth]; endX; endY;
                            // line; [objectColor]; posX; posY; [rotate]; [lineWidth]; lineLength;
                            else if (Label[i].type == _objectNames[lineObject])
                            {
                                if (cells.Length >= 7)
                                {
                                    if (cells[1] != "")
                                    {
                                        Label[i].fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        Label[i].fgColor = fgColor;
                                    }
                                    if (Label[i].bgColor == Label[i].fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out Label[i].posX);
                                    if (Label[i].posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = 0;
                                    }
                                    else if (Label[i].posX >= labelWidth)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = labelWidth - 1;
                                    }

                                    float.TryParse(cells[3], out Label[i].posY);
                                    if (Label[i].posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = 0;
                                    }
                                    else if (Label[i].posY >= labelHeight)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = labelHeight - 1;
                                    }

                                    float.TryParse(cells[5], out Label[i].lineWidth);
                                    if (Label[i].lineWidth < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect line width: " + Label[i].lineWidth.ToString());
                                        Label[i].lineWidth = 0;
                                    }

                                    if (cells[4] != "")
                                    {
                                        float.TryParse(cells[4], out Label[i].rotate);

                                        float.TryParse(cells[6], out Label[i].lineLength);
                                        if (Label[i].lineLength < 0)
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect line length: " + Label[i].lineLength.ToString());
                                            Label[i].lineLength = 0;
                                        }
                                        else if (Label[i].lineLength >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect line length: " + Label[i].lineLength.ToString());
                                            Label[i].lineLength = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                        }
                                    }
                                    else
                                    {
                                        float.TryParse(cells[6], out Label[i].width);
                                        if (Label[i].width < 0)
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                            Label[i].width = 0;
                                        }
                                        else if (Label[i].width >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                            Label[i].width = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                        }

                                        float.TryParse(cells[7], out Label[i].height);
                                        if (Label[i].height < 0)
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                            Label[i].height = 0;
                                        }
                                        else if (Label[i].height >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                            Label[i].height = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                        }

                                        Label[i].lineLength = -1;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // rectangle; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
                            else if (Label[i].type == _objectNames[rectangleObject])
                            {
                                if (cells.Length >= 9)
                                {
                                    if (cells[1] != "")
                                    {
                                        Label[i].fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        Label[i].fgColor = fgColor;
                                    }
                                    if (Label[i].bgColor == Label[i].fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out Label[i].posX);
                                    if (Label[i].posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = 0;
                                    }
                                    else if (Label[i].posX >= labelWidth)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = labelWidth - 1;
                                    }

                                    float.TryParse(cells[3], out Label[i].posY);
                                    if (Label[i].posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = 0;
                                    }
                                    else if (Label[i].posY >= labelHeight)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = labelHeight - 1;
                                    }

                                    float.TryParse(cells[4], out Label[i].rotate);

                                    float.TryParse(cells[5], out Label[i].lineWidth);
                                    if (Label[i].lineWidth < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect line width: " + Label[i].lineWidth.ToString());
                                        Label[i].lineWidth = 0;
                                    }

                                    float.TryParse(cells[6], out Label[i].width);
                                    if (Label[i].width < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                        Label[i].width = 0;
                                    }
                                    else if (Label[i].width >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                        Label[i].width = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                    }

                                    float.TryParse(cells[7], out Label[i].height);
                                    if (Label[i].height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                        Label[i].height = 0;
                                    }
                                    else if (Label[i].height >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                        Label[i].height = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                    }

                                    byte t = 0;
                                    byte.TryParse(cells[8], out t);
                                    Label[i].fill = (t > 0);
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                            // ellipse; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
                            else if (Label[i].type == _objectNames[ellipseObject])
                            {
                                if (cells.Length >= 9)
                                {
                                    if (cells[1] != "")
                                    {
                                        Label[i].fgColor = Color.FromName(cells[1]);
                                    }
                                    else
                                    {
                                        Label[i].fgColor = fgColor;
                                    }
                                    if (Label[i].bgColor == Label[i].fgColor)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                    }

                                    float.TryParse(cells[2], out Label[i].posX);
                                    if (Label[i].posX < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = 0;
                                    }
                                    else if (Label[i].posX >= labelWidth)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect X position: " + Label[i].posX.ToString());
                                        Label[i].posX = labelWidth - 1;
                                    }

                                    float.TryParse(cells[3], out Label[i].posY);
                                    if (Label[i].posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = 0;
                                    }
                                    else if (Label[i].posY >= labelHeight)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect Y position: " + Label[i].posY.ToString());
                                        Label[i].posY = labelHeight - 1;
                                    }

                                    float.TryParse(cells[4], out Label[i].rotate);

                                    float.TryParse(cells[5], out Label[i].lineWidth);
                                    if (Label[i].lineWidth < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect line width: " + Label[i].lineWidth.ToString());
                                        Label[i].lineWidth = 0;
                                    }

                                    float.TryParse(cells[6], out Label[i].width);
                                    if (Label[i].width < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                        Label[i].width = 0;
                                    }
                                    else if (Label[i].width >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].width.ToString());
                                        Label[i].width = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                    }

                                    float.TryParse(cells[7], out Label[i].height);
                                    if (Label[i].height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                        Label[i].height = 0;
                                    }
                                    else if (Label[i].height >= Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + Label[i].height.ToString());
                                        Label[i].height = (int)Math.Sqrt(labelWidth * labelWidth + labelHeight * labelHeight);
                                    }

                                    byte t = 0;
                                    byte.TryParse(cells[8], out t);
                                    Label[i].fill = (t > 0);
                                }
                                else
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                                }
                            }
                        }
                    }
                }
                button_importLabels.Enabled = true;
                generateLabel(-1);
                textBox_templateName.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\') + 1);
                //create colums and fill 1 row with default values
                for (int i = 1; i < objectsNum; i++)
                {
                    LabelsDatabase.Columns.Add((i - 1).ToString() + " " + Label[i].type);
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
            LabelBmp = new Bitmap(labelWidth, labelHeight, PixelFormat.Format32bppPArgb);
            if (checkBox_scale.Checked)
            {
                pictureBox_label.Dock = DockStyle.None;
                pictureBox_label.Width = labelWidth;
                pictureBox_label.Height = labelHeight;
            }
            else pictureBox_label.Dock = DockStyle.Fill;
            for (int i = 0; i < Label.Length; i++)
            {
                if (Label[i].type == "label")
                {
                    fillBackground(Label[i].bgColor);
                }
                else if (Label[i].type == "text")
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
                else if (Label[i].type == "picture")
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
                else if (Label[i].type == "barcode")
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
                else if (Label[i].type == "line")
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
                else if (Label[i].type == "rectangle")
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
                else if (Label[i].type == "ellipse")
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
                else MessageBox.Show("Incorrect object: " + Label[i].type);
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
            Bitmap bmp = new Bitmap(labelWidth, labelHeight);
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
            pictureBox_label.Width = labelWidth;
            pictureBox_label.Height = labelHeight;

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
                pictureBox_label.Width = labelWidth;
                pictureBox_label.Height = labelHeight;
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

    }
}
