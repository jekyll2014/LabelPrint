using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ZXing;
using ZXing.Common;

namespace LabelPrint
{
    public partial class Form1 : Form
    {
        private const int labelObject = 0;
        private const int textObject = 1;
        private const int pictureObject = 2;
        private const int barcodeObject = 3;
        private const int lineObject = 4;
        private const int rectangleObject = 5;
        private const int ellipseObject = 6;
        private string[] _objectNames = { "label", "text", "picture", "barcode", "line", "rectangle", "ellipse", };
        private string[] _textStyleNames = { "0=regular", "1=bold", "2=italic", "4=underline", "8=strikeout" };

        private struct template
        {
            public string name;
            public Color bgColor;
            public Color fgColor;
            public float dpi;
            public int codePage;
            public float posX;
            public float posY;
            public float rotate;
            public string content;
            public float width;
            public float height;
            public bool transparent;
            public int BCformat;
            public float fontSize;
            public byte fontStyle;
            public string fontName;
            public string feature;
            public float lineLength;
            public float lineWidth;
        }

        private List<template> Label = new List<template>();
        private DataTable LabelsDatabase = new DataTable();
        private int objectsNum = 0;
        private List<string> bcFeatures = new List<string> { "AZTEC_LAYERS", "ERROR_CORRECTION", "MARGIN", "PDF417_ASPECT_RATIO", "QR_VERSION" };
        private List<int> BarCodeTypes = new List<int> {
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

        private float mult = 1;

        // measurement units constants [pix, mm, cm, "]
        private float[] units = { 1, 1, 1, 1 };
        private int pagesFrom = 0;
        private int pagesTo = 0;
        private bool _templateChanged = false;
        private bool cmdLinePrint = false;
        private string printerName = "";
        private Bitmap LabelBmp = new Bitmap(1, 1, PixelFormat.Format32bppPArgb);
        private Rectangle currentObject = new Rectangle();
        private Color _borderColor = Color.Black;
        private string path = "";

        public Form1(string[] cmdLine)
        {
            InitializeComponent();
            if (cmdLine.Length >= 1)
            {
                tabControl1.SelectedIndexChanged -= new EventHandler(tabControl1_SelectedIndexChanged);
                listBox_objects.SelectedIndexChanged -= new EventHandler(listBox_objects_SelectedIndexChanged);
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
                if (Application.MessageLoop)
                {
                    // WinForms app
                    Application.Exit();
                }
                else
                {
                    // Console app
                    Environment.Exit(1);
                }
            }
            comboBox_object.Items.AddRange(_objectNames);

            template init_label = new template();
            init_label.name = _objectNames[labelObject];
            init_label.bgColor = Color.White;
            init_label.fgColor = Color.Black;
            init_label.width = 1;
            init_label.height = 1;
            init_label.dpi = 200;
            init_label.codePage = 65001;
            Label.Add(init_label);

            listBox_objects.Items.AddRange(getObjectsList());
            listBox_objects.SelectedIndex = 0;
            comboBox_units.SelectedIndex = 0;
            textBox_dpi.Text = Label[0].dpi.ToString("F4");

            comboBox_encoding.Items.Clear();
            comboBox_encoding.Items.AddRange(getEncodingList());
            string str = Properties.Settings.Default.CodePage.ToString();
            for (int i = 0; i < comboBox_encoding.Items.Count; i++)
            {
                if (comboBox_encoding.Items[i].ToString().StartsWith(str))
                {
                    comboBox_encoding.SelectedIndex = i;
                    break;
                }
            }
            textBox_dpi_Leave(this, EventArgs.Empty);
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

        private Rectangle drawText(Bitmap img, Color fgC, float posX, float posY, string text, string fontName, float fontSize, float rotateDeg = 0, FontStyle fontStyle = FontStyle.Regular)
        {
            Rectangle size = new Rectangle(0, 0, 0, 0);
            //check if font supports all the options
            Font textFont;
            try
            {
                textFont = new Font(fontName, fontSize, fontStyle, GraphicsUnit.Pixel); //creates new font
            }
            catch (Exception ex)
            {
                MessageBox.Show("Font error: " + ex.Message);
                return size;
            }
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
            SizeF stringSize = g.MeasureString(text, textFont);
            size.X = (int)posX;
            size.Y = (int)posY;
            size.Width = (int)stringSize.Width + 1;
            size.Height = (int)fontSize + 1;
            return size;
        }

        private Rectangle drawPicture(Bitmap img, Color fgC, float posX, float posY, string fileName, float rotateDeg = 0, float width = 0, float height = 0, bool makeTransparent = true)
        {
            Rectangle size = new Rectangle(0, 0, 0, 0);
            Bitmap newPicture = new Bitmap(1, 1);
            try
            {
                newPicture = new Bitmap(@fileName, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file: " + fileName + " : " + ex.Message);
                return size;
            }
            if (width == 0) width = newPicture.Width; //-V3024
            if (height == 0) height = newPicture.Height; //-V3024
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
            size.X = (int)posX;
            size.Y = (int)posY;
            size.Width = (int)width + 1;
            size.Height = (int)height + 1;
            return size;
        }

        private Rectangle drawBarcode(Bitmap img, Color bgC, Color fgC, float posX, float posY, float width, float height, string BCdata, BarcodeFormat bcFormat, float rotateDeg = 0, string addFeature = "", bool makeTransparent = true)
        {
            Rectangle size = new Rectangle(0, 0, 0, 0);
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
            Bitmap newPicture;
            try
            {
                newPicture = barCode.Write(BCdata);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Barcode generation error: " + ex.Message);
                return size;
            }
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
            size.X = (int)posX;
            size.Y = (int)posY;
            size.Width = (int)width + 1;
            size.Height = (int)height + 1;
            return size;
        }

        private Rectangle drawLineCoord(Bitmap img, Color fgC, float posX, float posY, float endX, float endY, float lineWidth)
        {
            Rectangle size = new Rectangle(0, 0, 0, 0);
            Pen p = new Pen(fgC, lineWidth);
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawLine(p, posX, posY, endX, endY);
            size.X = (int)posX;
            size.Y = (int)posY;
            size.Width = (int)(endX - posX) + 1;
            size.Height = (int)(endY - posY) + 1;
            if (size.Width >= 0)
            {
                size.X = size.X - (int)(lineWidth / 2) - 1;
                size.Width = size.Width + (int)lineWidth + 1;
            }
            else
            {
                size.X = size.X + (int)(lineWidth / 2) + 1;
                size.Width = size.Width - (int)lineWidth - 1;
                size.X = size.X + size.Width;
                size.Width = -size.Width;
            }
            if (size.Height >= 0)
            {
                size.Y = size.Y - (int)(lineWidth / 2) - 1;
                size.Height = size.Height + (int)lineWidth + 1;
            }
            else
            {
                size.Y = size.Y + (int)(lineWidth / 2) + 1;
                size.Height = size.Height - (int)lineWidth - 1;
                size.Y = size.Y + size.Height;
                size.Height = -size.Height;
            }
            return size;
        }

        private Rectangle drawLineLength(Bitmap img, Color fgC, float posX, float posY, float length, float rotateDeg, float lineWidth)
        {
            Rectangle size = new Rectangle(0, 0, 0, 0);
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
            size.X = (int)posX;
            size.Y = (int)posY;
            size.Width = (int)(length * Math.Cos(rotateDeg * Math.PI / 180)) + 1;
            size.Height = (int)(length * Math.Sin(rotateDeg * Math.PI / 180)) + 1;
            if (size.Width >= 0)
            {
                size.X = size.X - (int)(lineWidth / 2) - 1;
                size.Width = size.Width + (int)lineWidth + 1;
            }
            else
            {
                size.X = size.X + (int)(lineWidth / 2) + 1;
                size.Width = size.Width - (int)lineWidth - 1;
                size.X = size.X + size.Width;
                size.Width = -size.Width;
            }
            if (size.Height >= 0)
            {
                size.Y = size.Y - (int)(lineWidth / 2) - 1;
                size.Height = size.Height + (int)lineWidth + 1;
            }
            else
            {
                size.Y = size.Y + (int)(lineWidth / 2) + 1;
                size.Height = size.Height - (int)lineWidth - 1;
                size.Y = size.Y + size.Height;
                size.Height = -size.Height;
            }
            return size;
        }

        private Rectangle drawRectangle(Bitmap img, Color fgC, float posX, float posY, float width, float height, float rotateDeg, float lineWidth, bool fill)
        {
            Rectangle size = new Rectangle(0, 0, 0, 0);
            Pen p = new Pen(fgC, lineWidth);
            p.Alignment = PenAlignment.Inset;
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
            size.X = (int)posX;
            size.Y = (int)posY;
            size.Width = (int)width + 1;
            size.Height = (int)height + 1;
            return size;
        }

        private Rectangle drawEllipse(Bitmap img, Color fgC, float posX, float posY, float width, float height, float rotateDeg, float lineWidth, bool fill)
        {
            Rectangle size = new Rectangle(0, 0, 0, 0);
            Pen p = new Pen(fgC, lineWidth);
            p.Alignment = PenAlignment.Inset;
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
            size.X = (int)posX;
            size.Y = (int)posY;
            size.Width = (int)width + 1;
            size.Height = (int)height + 1;
            return size;
        }

        private void ReadCsv(string fileName, DataTable table, bool createColumnsNames = false)
        {
            table.Clear();
            table.Rows.Clear();
            List<string> inputStr = new List<string>();
            try
            {
                inputStr.AddRange(File.ReadAllLines(fileName, Encoding.GetEncoding(Properties.Settings.Default.CodePage)));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file:" + fileName + " : " + ex.Message);
                return;
            }

            //read headers
            int n = 0;
            if (createColumnsNames == true)
            {
                table.Columns.Clear();
                //create and count columns and read headers
                if (inputStr[n].Length != 0)
                {
                    string[] cells = inputStr[0].ToString().Split(Properties.Settings.Default.CSVdelimiter);
                    for (int i = 0; i < cells.Length - 1; i++)
                    {
                        table.Columns.Add(cells[i]);
                    }
                }
                n++;
            }
            else
            {
                //create 1st row and count columns
                if (inputStr[0].Length != 0)
                {
                    string[] cells = inputStr[0].ToString().Split(Properties.Settings.Default.CSVdelimiter);
                    DataRow row = table.NewRow();
                    for (int i = 0; i < cells.Length - 1; i++)
                    {
                        string tmp1 = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').Trim();
                        if (tmp1 != "") row[i] = tmp1;
                    }
                    table.Rows.Add(row);
                }
            }

            //read CSV content string by string
            for (int i1 = 1; i1 < inputStr.Count; i1++)
            {
                if (inputStr[i1].ToString().Replace(Properties.Settings.Default.CSVdelimiter, ' ').Trim().TrimStart('\r').TrimStart('\n').TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').TrimEnd('\r').TrimEnd('\n') != "")
                {
                    string[] cells = inputStr[i1].ToString().Split(Properties.Settings.Default.CSVdelimiter);
                    DataRow row = table.NewRow();
                    for (int i = 0; i < cells.Length - 1; i++)
                    {
                        //row[i] = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').Trim();
                        string tmp1 = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').Trim();
                        if (tmp1 != "") row[i] = tmp1;
                    }
                    table.Rows.Add(row);
                }
            }
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
                            if (Label[i + 1].name == _objectNames[pictureObject] && !File.Exists(path + row.Cells[i].Value.ToString()))
                            {
                                MessageBox.Show("[Line " + (i + 1).ToString() + "] File not exist: " + path + row.Cells[i].Value.ToString());
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
                string[] inputStr = File.ReadAllLines(openFileDialog1.FileName, Encoding.GetEncoding(Properties.Settings.Default.CodePage));
                Label.Clear();
                for (int i = 0; i < inputStr.Length; i++)
                {
                    if (inputStr[i].Trim() != "")
                    {
                        List<string> cells = new List<string>();
                        foreach (string str in inputStr[i].Split(Properties.Settings.Default.CSVdelimiter))
                        {
                            cells.Add(str.Trim());
                        }
                        if (cells[cells.Count - 1] == "") cells.RemoveAt(cells.Count - 1);
                        if (cells.Count >= 7)
                        {
                            template templ = new template();
                            if (i == 0 && cells[0] != _objectNames[labelObject])
                            {
                                MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                                return;
                            }
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

                                templ.dpi = 0;
                                float.TryParse(cells[5], out templ.dpi);
                                if (templ.dpi == 0) float.TryParse(textBox_dpi.Text, out templ.dpi); //-V3024
                                else if (templ.dpi < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect resolution: " + templ.dpi.ToString());
                                    templ.height = 1;
                                }
                                textBox_dpi.Text = templ.dpi.ToString("F4");

                                int.TryParse(cells[6], out templ.codePage);
                                if (templ.codePage == 0) templ.codePage = Properties.Settings.Default.CodePage;
                                else if (templ.codePage < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect codepage: " + templ.codePage.ToString());
                                    templ.height = 1;
                                }
                                for (int j = 0; j < comboBox_encoding.Items.Count; j++)
                                {
                                    if (comboBox_encoding.Items[j].ToString().StartsWith(templ.codePage.ToString()))
                                    {
                                        comboBox_encoding.SelectedIndex = j;
                                        break;
                                    }
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
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
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

                                    float.TryParse(cells[7], out templ.fontSize);

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
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                        templ.posY = (int)Label[0].height - 1;
                                    }

                                    float.TryParse(cells[4], out templ.rotate);

                                    templ.content = cells[5];
                                    if (!File.Exists(path + templ.content))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] File not exist: " + path + templ.content);
                                    }

                                    float.TryParse(cells[6], out templ.width);
                                    if (templ.width < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = 0;
                                    }
                                    /*else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }*/

                                    float.TryParse(cells[7], out templ.height);
                                    if (templ.height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = 0;
                                    }
                                    /*else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }*/

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
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[4], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
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
                                    /*else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }*/

                                    float.TryParse(cells[8], out templ.height);
                                    if (templ.height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = 0;
                                    }
                                    /*else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }*/

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
                                if (cells.Count >= 7) //-V3022
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
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
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
                                        /*else if (templ.lineLength >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect line length: " + templ.lineLength.ToString());
                                            templ.lineLength = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                        }*/
                                    }
                                    else
                                    {
                                        float.TryParse(cells[6], out templ.width);
                                        if (templ.width < 0)
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                            templ.width = 0;
                                        }
                                        /*else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                            templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                        }*/

                                        float.TryParse(cells[7], out templ.height);
                                        if (templ.height < 0)
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                            templ.height = 0;
                                        }
                                        /*else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                        {
                                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                            templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                        }*/

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
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
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
                                    /*else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }*/

                                    float.TryParse(cells[7], out templ.height);
                                    if (templ.height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = 0;
                                    }
                                    /*else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }*/

                                    byte t = 0;
                                    byte.TryParse(cells[8], out t);
                                    templ.transparent = (t > 0);
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
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = 0;
                                    }
                                    else if (templ.posX >= (int)Label[0].width)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                        templ.posX = (int)Label[0].width - 1;
                                    }

                                    float.TryParse(cells[3], out templ.posY);
                                    if (templ.posY < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                        templ.posY = 0;
                                    }
                                    else if (templ.posY >= (int)Label[0].height)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
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
                                    /*else if (templ.width >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }*/

                                    float.TryParse(cells[7], out templ.height);
                                    if (templ.height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = 0;
                                    }
                                    /*else if (templ.height >= Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = (int)Math.Sqrt((int)Label[0].width * (int)Label[0].width + (int)Label[0].height * (int)Label[0].height);
                                    }*/

                                    byte t = 0;
                                    byte.TryParse(cells[8], out t);
                                    templ.transparent = (t > 0);
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
                textBox_dpi_Leave(this, EventArgs.Empty);
                button_importLabels.Enabled = true;
                generateLabel(-1);
                textBox_templateName.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\') + 1);
                //save path to template
                path = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.LastIndexOf('\\') + 1);
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
                tabControl1_SelectedIndexChanged(this, EventArgs.Empty);
            }
        }

        private void generateLabel(int gridLine)
        {
            if (Label[0].name != _objectNames[labelObject])
            {
                MessageBox.Show("1st object must be \"label\"");
                return;
            }
            if (checkBox_scale.Checked)
            {
                pictureBox_label.Dock = DockStyle.None;
                pictureBox_label.Width = (int)Label[0].width;
                pictureBox_label.Height = (int)Label[0].height;
            }
            else pictureBox_label.Dock = DockStyle.Fill;

            Rectangle r = new Rectangle();
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
                    float fontSize = Label[i].fontSize;
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    string content = Label[i].content;
                    FontStyle fontStyle = (FontStyle)Label[i].fontStyle;
                    string tmp = "";
                    if (gridLine > -1) tmp = dataGridView_labels.Rows[gridLine].Cells[i - 1].Value.ToString();
                    if (tmp != "") content = tmp;
                    r = drawText(LabelBmp, Label[i].fgColor, posX, posY, content, fontname, fontSize, rotate, fontStyle);
                }
                else if (Label[i].name == _objectNames[pictureObject])
                {
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    string content = path + Label[i].content;
                    float width = Label[i].width;
                    float height = Label[i].height;
                    bool transparent = Label[i].transparent;
                    string tmp = "";
                    if (gridLine > -1) tmp = path + dataGridView_labels.Rows[gridLine].Cells[i - 1].Value.ToString();
                    if (tmp != "") content = tmp;
                    r = drawPicture(LabelBmp, Label[i].fgColor, posX, posY, content, rotate, width, height, transparent);
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
                    r = drawBarcode(LabelBmp, Label[i].bgColor, Label[i].fgColor, posX, posY, width, height, content, BCformat, rotate, feature, transparent);
                }
                else if (Label[i].name == _objectNames[lineObject])
                {
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    float lineWidth = Label[i].lineWidth;
                    if (Label[i].lineLength == -1) //-V3024
                    {
                        float endX = Label[i].width;
                        float endY = Label[i].height;
                        r = drawLineCoord(LabelBmp, Label[i].fgColor, posX, posY, endX, endY, lineWidth);
                    }
                    else
                    {
                        float length = Label[i].lineLength;
                        r = drawLineLength(LabelBmp, Label[i].fgColor, posX, posY, length, rotate, lineWidth);
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
                    bool fill = !Label[i].transparent;
                    r = drawRectangle(LabelBmp, Label[i].fgColor, posX, posY, width, height, rotate, lineWidth, fill);
                }
                else if (Label[i].name == _objectNames[ellipseObject])
                {
                    float posX = Label[i].posX;
                    float posY = Label[i].posY;
                    float rotate = Label[i].rotate;
                    float lineWidth = Label[i].lineWidth;
                    float width = Label[i].width;
                    float height = Label[i].height;
                    bool fill = !Label[i].transparent;
                    r = drawEllipse(LabelBmp, Label[i].fgColor, posX, posY, width, height, rotate, lineWidth, fill);
                }
                else MessageBox.Show("Incorrect object: " + Label[i].name);
                if (i == listBox_objects.SelectedIndex) currentObject = r;
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
            if (tabControl1.SelectedIndex == 0)
            {
                timer1.Enabled = false;
                if (_templateChanged)
                {
                    dataGridView_labels.SelectionChanged -= new EventHandler(dataGridView_labels_SelectionChanged);
                    dataGridView_labels.DataSource = null;
                    textBox_labelsName.Clear();
                    LabelsDatabase.Clear();
                    LabelsDatabase.Rows.Clear();
                    List<string> inputStr = new List<string>();
                    char div = Properties.Settings.Default.CSVdelimiter;
                    //create column headers
                    LabelsDatabase.Columns.Clear();
                    //create and count columns and read headers
                    for (int i = 1; i < Label.Count; i++)
                    {
                        LabelsDatabase.Columns.Add(i.ToString() + " " + Label[i].name);
                    }

                    //create 1st row and count columns
                    DataRow r = LabelsDatabase.NewRow();
                    for (int i = 1; i < Label.Count; i++)
                    {
                        r[i - 1] = Label[i].content;
                    }
                    LabelsDatabase.Rows.Add(r);

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
                    dataGridView_labels.CurrentCell = dataGridView_labels.Rows[0].Cells[0];
                    dataGridView_labels.Rows[0].Selected = true;
                    generateLabel(-1);
                    dataGridView_labels.SelectionChanged += new EventHandler(dataGridView_labels_SelectionChanged);
                    _templateChanged = false;
                }
                generateLabel(-1);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(getObjectsList());
                listBox_objects.SelectedIndex = 0;
                timer1.Enabled = true;
            }
        }

        private void listBox_objects_SelectedIndexChanged(object sender, EventArgs e)
        {
            showObject(listBox_objects.SelectedIndex);
        }

        private void button_apply_Click(object sender, EventArgs e)
        {
            listBox_objects.SelectedIndexChanged -= new EventHandler(listBox_objects_SelectedIndexChanged);
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
            listBox_objects.Items.Clear();
            listBox_objects.Items.AddRange(getObjectsList());
            listBox_objects.SelectedIndex = n;
            _templateChanged = true;
            showObject(listBox_objects.SelectedIndex);
            listBox_objects.SelectedIndexChanged += new EventHandler(listBox_objects_SelectedIndexChanged);
        }

        private void button_delete_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < listBox_objects.Items.Count - 1 && listBox_objects.SelectedIndex > 0)
            {
                int n = listBox_objects.SelectedIndex;
                Label.RemoveAt(n);
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(getObjectsList());
                listBox_objects.SelectedIndex = n;
                _templateChanged = true;
                showObject(listBox_objects.SelectedIndex);
                generateLabel(-1);
            }
        }

        private void button_up_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < listBox_objects.Items.Count - 1 && listBox_objects.SelectedIndex > 1)
            {
                int n = listBox_objects.SelectedIndex;
                template templ = Label[n];
                Label.RemoveAt(n);
                Label.Insert(n - 1, templ);
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(getObjectsList());
                listBox_objects.SelectedIndex = n - 1;
                _templateChanged = true;
                showObject(listBox_objects.SelectedIndex);
                generateLabel(-1);
            }
        }

        private void button_down_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < listBox_objects.Items.Count - 2 && listBox_objects.SelectedIndex > 0)
            {
                int n = listBox_objects.SelectedIndex;
                template templ = Label[n];
                Label.RemoveAt(n);
                Label.Insert(n + 1, templ);
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(getObjectsList());
                listBox_objects.SelectedIndex = n + 1;
                _templateChanged = true;
                showObject(listBox_objects.SelectedIndex);
                generateLabel(-1);
            }
        }

        private string[] getObjectsList()
        {
            List<string> objectList = new List<string>();
            foreach (template t in Label)
            {
                objectList.Add(t.name);
            }
            objectList.Add("");
            return objectList.ToArray();
        }

        private string[] getColorList()
        {
            List<string> colorList = new List<string>();
            foreach (Color c in new ColorConverter().GetStandardValues())
            {
                colorList.Add(c.Name);
            }
            return colorList.ToArray();
        }

        private string[] getFontList()
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

        private string[] getBarcodeList()
        {
            List<string> barcodeList = new List<string>();
            foreach (BarcodeFormat b in BarCodeTypes)
            {
                barcodeList.Add(((int)b).ToString() + "=" + b.ToString());
            }
            return barcodeList.ToArray();
        }

        private string[] getEncodingList()
        {
            List<string> encodeList = new List<string>();

            EncodingInfo[] codePages = Encoding.GetEncodings();
            foreach (EncodingInfo codePage in codePages)
            {
                encodeList.Add(codePage.CodePage.ToString() + "=" + "[" + codePage.Name + "]" + codePage.DisplayName);
            }
            return encodeList.ToArray();
        }

        private void showObject(int n)
        {
            clearFields();
            comboBox_object.SelectedItem = listBox_objects.SelectedItem.ToString();
            string str = "";
            //new object
            if (n >= Label.Count)
            {
                comboBox_object.Enabled = true;
            }
            // label; [bgColor]; [objectColor]; width; height;
            else if (Label[n].name == _objectNames[labelObject])
            {
                label_backgroundColor.Text = "Background color";
                comboBox_backgroundColor.Enabled = true;
                comboBox_backgroundColor.Items.Clear();

                comboBox_backgroundColor.Items.AddRange(getColorList());
                str = Label[n].bgColor.Name.ToString();
                for (int i = 0; i < comboBox_backgroundColor.Items.Count; i++)
                {
                    if (comboBox_backgroundColor.Items[i].ToString() == str)
                    {
                        comboBox_backgroundColor.SelectedIndex = i;
                        i = comboBox_backgroundColor.Items.Count;
                    }
                }

                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Enabled = true;
                comboBox_objectColor.Items.Clear();
                comboBox_objectColor.Items.Add("");
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[n].fgColor.Name.ToString();
                for (int i = 0; i < comboBox_objectColor.Items.Count; i++)
                {
                    if (comboBox_objectColor.Items[i].ToString() == str)
                    {
                        comboBox_objectColor.SelectedIndex = i;
                        i = comboBox_objectColor.Items.Count;
                    }
                }

                label_width.Text = "Width";
                textBox_width.Enabled = true;
                textBox_width.Text = (Label[n].width / mult).ToString("F4");

                label_height.Text = "Height";
                textBox_height.Enabled = true;
                textBox_height.Text = (Label[n].height / mult).ToString("F4");
            }

            // text; [objectColor]; posX; posY; [rotate]; [default_text]; fontName; fontSize; [fontStyle];
            else if (Label[n].name == _objectNames[textObject])
            {
                comboBox_objectColor.Enabled = true;
                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[n].fgColor.Name.ToString();
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
                textBox_posX.Text = (Label[n].posX / mult).ToString("F4");

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = (Label[n].posY / mult).ToString("F4");

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[n].rotate.ToString();

                textBox_content.Enabled = true;
                label_content.Text = "Data string";
                textBox_content.Text = Label[n].content;

                comboBox_fontName.Enabled = true;
                label_fontName.Text = "Font";
                comboBox_fontName.Items.Clear();
                comboBox_fontName.Items.AddRange(getFontList());
                comboBox_fontName.SelectedIndex = 0;
                str = Label[n].fontName.ToString();
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
                comboBox_fontStyle.SelectedIndex = 0;
                str = Label[n].fontStyle.ToString();
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
                textBox_fontSize.Text = (Label[n].fontSize / mult).ToString("F4");
            }

            // picture; [objectColor]; posX; posY; [rotate]; [default_file]; [width]; [height]; [transparent];
            else if (Label[n].name == _objectNames[pictureObject])
            {
                comboBox_objectColor.Enabled = true;
                label_objectColor.Text = "Transparent color";
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[n].fgColor.Name.ToString();
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
                textBox_posX.Text = (Label[n].posX / mult).ToString("F4");

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = (Label[n].posY / mult).ToString("F4");

                textBox_width.Enabled = true;
                label_width.Text = "Width";
                textBox_width.Text = (Label[n].width / mult).ToString("F4");

                textBox_height.Enabled = true;
                label_height.Text = "Height";
                textBox_height.Text = (Label[n].height / mult).ToString("F4");

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[n].rotate.ToString();

                textBox_content.Enabled = true;
                label_content.Text = "Picture file";
                textBox_content.Text = Label[n].content;

                checkBox_fill.Enabled = true;
                checkBox_fill.Text = "Transparent";
                checkBox_fill.Checked = Label[n].transparent;
            }

            // barcode; [bgColor]; [objectColor]; posX; posY; [rotate]; [default_data]; width; height; bcFormat; [transparent]; [additional_features]
            else if (Label[n].name == _objectNames[barcodeObject])
            {
                comboBox_backgroundColor.Enabled = true;
                label_backgroundColor.Text = "Background color";
                comboBox_backgroundColor.Items.AddRange(getColorList());
                str = Label[n].bgColor.Name.ToString();
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
                str = Label[n].fgColor.Name.ToString();
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
                textBox_posX.Text = (Label[n].posX / mult).ToString("F4");

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = (Label[n].posY / mult).ToString("F4");

                textBox_width.Enabled = true;
                label_width.Text = "Width";
                textBox_width.Text = (Label[n].width / mult).ToString("F4");

                textBox_height.Enabled = true;
                label_height.Text = "Height";
                textBox_height.Text = (Label[n].height / mult).ToString("F4");

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[n].rotate.ToString();

                textBox_content.Enabled = true;
                label_content.Text = "Data string";
                textBox_content.Text = Label[n].content;

                comboBox_fontName.Enabled = true;
                label_fontName.Text = "Barcode type";
                comboBox_fontName.Items.Clear();
                comboBox_fontName.Items.AddRange(getBarcodeList());
                comboBox_fontName.SelectedIndex = 0;
                str = Label[n].BCformat.ToString() + "=" + ((BarcodeFormat)Label[n].BCformat).ToString();
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
                comboBox_fontStyle.SelectedIndex = 0;
                str = Label[n].feature;
                for (int i = 0; i < comboBox_fontStyle.Items.Count; i++)
                {
                    if (str.StartsWith(comboBox_fontStyle.Items[i].ToString()))
                    {
                        comboBox_fontStyle.SelectedIndex = i;
                    }
                }

                textBox_fontSize.Enabled = true;
                label_fontSize.Text = "Feature value";
                textBox_fontSize.Text = Label[n].feature.Substring(Label[n].feature.IndexOf('=') + 1);

                checkBox_fill.Enabled = true;
                checkBox_fill.Text = "Transparent";
                checkBox_fill.Checked = Label[n].transparent;
            }

            // line; [objectColor]; posX; posY; --- ; [lineWidth]; endX; endY; (lineLength = -1)
            // line; [objectColor]; posX; posY; [rotate]; [lineWidth]; lineLength;
            else if (Label[n].name == _objectNames[lineObject])
            {
                comboBox_objectColor.Enabled = true;
                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[n].fgColor.Name.ToString();
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
                textBox_posX.Text = (Label[n].posX / mult).ToString("F4");

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = (Label[n].posY / mult).ToString("F4");

                textBox_fontSize.Enabled = true;
                label_fontSize.Text = "Line width";
                textBox_fontSize.Text = (Label[n].lineWidth / mult).ToString("F4");

                textBox_width.Enabled = true;
                label_width.Text = "endX (empty to use length)";
                textBox_width.Text = (Label[n].width / mult).ToString("F4");

                textBox_height.Enabled = true;
                label_height.Text = "endY (empty to use length)";
                textBox_height.Text = (Label[n].height / mult).ToString("F4");
                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate (empty to use length)";
                textBox_rotate.Text = Label[n].rotate.ToString();

                textBox_content.Enabled = true;
                label_content.Text = "Line length (empty to use coordinates)";
                if (Label[n].lineLength != -1) textBox_content.Text = (Label[n].lineLength / mult).ToString("F4"); //-V3024
            }

            // rectangle; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
            else if (Label[n].name == _objectNames[rectangleObject])
            {
                comboBox_objectColor.Enabled = true;
                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[n].fgColor.Name.ToString();
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
                textBox_posX.Text = (Label[n].posX / mult).ToString("F4");

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = (Label[n].posY / mult).ToString("F4");

                textBox_width.Enabled = true;
                label_width.Text = "Width";
                textBox_width.Text = (Label[n].width / mult).ToString("F4");

                textBox_height.Enabled = true;
                label_height.Text = "Height";
                textBox_height.Text = (Label[n].height / mult).ToString("F4");

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[n].rotate.ToString();

                textBox_fontSize.Enabled = true;
                label_fontSize.Text = "Line width";
                textBox_fontSize.Text = (Label[n].lineWidth / mult).ToString("F4");

                checkBox_fill.Enabled = true;
                checkBox_fill.Text = "Fill with objectColor";
                checkBox_fill.Checked = !Label[n].transparent;
            }

            // ellipse; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
            else if (Label[n].name == _objectNames[ellipseObject])
            {
                label_objectColor.Text = "Default object color";
                comboBox_objectColor.Enabled = true;
                comboBox_objectColor.Items.AddRange(getColorList());
                str = Label[n].fgColor.Name.ToString();
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
                textBox_posX.Text = (Label[n].posX / mult).ToString("F4");

                textBox_posY.Enabled = true;
                label_posY.Text = "posY";
                textBox_posY.Text = (Label[n].posY / mult).ToString("F4");

                textBox_width.Enabled = true;
                label_width.Text = "Width";
                textBox_width.Text = (Label[n].width / mult).ToString("F4");

                textBox_height.Enabled = true;
                label_height.Text = "Height";
                textBox_height.Text = (Label[n].height / mult).ToString("F4");

                textBox_rotate.Enabled = true;
                label_rotate.Text = "Rotate";
                textBox_rotate.Text = Label[n].rotate.ToString();

                textBox_fontSize.Enabled = true;
                label_fontSize.Text = "Line width";
                textBox_fontSize.Text = (Label[n].lineWidth / mult).ToString("F4");

                checkBox_fill.Enabled = true;
                checkBox_fill.Text = "Fill with objectColor";
                checkBox_fill.Checked = !Label[n].transparent;
            }

            generateLabel(-1);
        }

        private template collectObject()
        {
            template templ = new template();

            templ.name = comboBox_object.SelectedItem.ToString();

            if (listBox_objects.SelectedIndex == listBox_objects.Items.Count - 1)
            {
                if (Label[0].name == _objectNames[labelObject])
                {
                    templ.bgColor = Label[0].bgColor;
                    templ.fgColor = Label[0].fgColor;
                }
                else
                {
                    templ.bgColor = Color.White;
                    templ.fgColor = Color.Black;
                }
                templ.posX = 0;
                templ.posY = 0;
                templ.dpi = 1;
                templ.dpi = Properties.Settings.Default.CodePage;
                templ.rotate = 0;
                templ.content = "";
                templ.width = 1;
                templ.height = 1;
                templ.transparent = true;
                templ.BCformat = 1;
                templ.fontSize = 10;
                templ.fontStyle = 0;
                templ.fontName = "Arial";
                templ.feature = "";
                templ.lineLength = 1;
                templ.lineWidth = 1;
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
                templ.width = f * mult;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f * mult;
            }

            // text; [objectColor]; posX; posY; [rotate]; [default_text]; fontName; fontSize; [fontStyle];
            else if (templ.name == _objectNames[textObject])
            {
                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.bgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f * mult;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f * mult;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                templ.content = textBox_content.Text;

                float b = 0;
                textBox_fontSize.Text = Evaluate(textBox_fontSize.Text);
                float.TryParse(textBox_fontSize.Text, out b);
                templ.fontSize = b * mult;

                byte s = 0;
                byte.TryParse(comboBox_fontStyle.SelectedItem.ToString().Substring(0, 1), out s);
                templ.fontStyle = s;

                templ.fontName = comboBox_fontName.SelectedItem.ToString();
            }

            // picture; [objectColor]; posX; posY; [rotate]; [default_file]; [width]; [height]; [transparent];
            else if (templ.name == _objectNames[pictureObject])
            {
                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.bgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f * mult;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f * mult;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                templ.content = textBox_content.Text;

                float.TryParse(textBox_width.Text, out f);
                templ.width = f * mult;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f * mult;

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
                templ.posX = f * mult;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f * mult;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                templ.content = textBox_content.Text;

                float.TryParse(textBox_width.Text, out f);
                templ.width = f * mult;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f * mult;

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
                templ.posX = f * mult;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f * mult;

                textBox_fontSize.Text = Evaluate(textBox_fontSize.Text);
                float.TryParse(textBox_fontSize.Text, out f);
                templ.lineWidth = f * mult;

                if ((textBox_rotate.Text == "" || textBox_content.Text == "" || templ.lineLength == -1) && textBox_width.Text != "" && textBox_height.Text != "") //-V3024
                {
                    float.TryParse(textBox_width.Text, out f);
                    templ.width = f * mult;

                    float.TryParse(textBox_height.Text, out f);
                    templ.height = f * mult;

                    templ.lineLength = -1;
                }
                else if (textBox_rotate.Text != "" && textBox_content.Text != "")
                {
                    float.TryParse(textBox_rotate.Text, out f);
                    templ.rotate = f;

                    textBox_content.Text = Evaluate(textBox_content.Text);
                    float.TryParse(textBox_content.Text, out f);
                    templ.lineLength = f * mult;
                }
            }

            // rectangle; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
            else if (templ.name == _objectNames[rectangleObject])
            {
                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f * mult;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f * mult;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                float.TryParse(textBox_width.Text, out f);
                templ.width = f * mult;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f * mult;

                textBox_fontSize.Text = Evaluate(textBox_fontSize.Text);
                float.TryParse(textBox_fontSize.Text, out f);
                templ.lineWidth = f * mult;

                templ.transparent = !checkBox_fill.Checked;
            }

            // ellipse; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
            else if (templ.name == _objectNames[ellipseObject])
            {
                if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                float f = 0;
                float.TryParse(textBox_posX.Text, out f);
                templ.posX = f * mult;

                float.TryParse(textBox_posY.Text, out f);
                templ.posY = f * mult;

                float.TryParse(textBox_rotate.Text, out f);
                templ.rotate = f;

                float.TryParse(textBox_width.Text, out f);
                templ.width = f * mult;

                float.TryParse(textBox_height.Text, out f);
                templ.height = f * mult;

                textBox_fontSize.Text = Evaluate(textBox_fontSize.Text);
                float.TryParse(textBox_fontSize.Text, out f);
                templ.lineWidth = f * mult;

                templ.transparent = !checkBox_fill.Checked;
            }
            return templ;
        }

        private void clearFields()
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
            comboBox_fontName.SelectedIndex = 0;
            textBox_fontSize.Enabled = false;

            checkBox_fill.Enabled = false;
            checkBox_fill.Text = "Transparent";

        }

        private void button_save_Click(object sender, EventArgs e)
        {
            SaveFileDialog1.Title = "Save template as .CSV...";
            SaveFileDialog1.DefaultExt = "csv";
            SaveFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
            if (textBox_templateName.Text == "") SaveFileDialog1.FileName = "template_" + DateTime.Today.ToShortDateString().Replace("/", "_") + ".csv";
            else SaveFileDialog1.FileName = textBox_templateName.Text;
            SaveFileDialog1.ShowDialog();
        }

        private void button_saveLabel_Click(object sender, EventArgs e)
        {
            SaveFileDialog1.Title = "Save label data as .CSV...";
            SaveFileDialog1.DefaultExt = "csv";
            SaveFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
            SaveFileDialog1.FileName = "label_" + textBox_templateName.Text;
            SaveFileDialog1.ShowDialog();
        }

        private void saveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StringBuilder output = new StringBuilder();
            char div = Properties.Settings.Default.CSVdelimiter;
            if (SaveFileDialog1.Title == "Save template as .CSV...")
            {
                for (int i = 0; i < Label.Count; i++)
                {
                    // label; 1 [bgColor]; 2 [objectColor]; 3 width; 4 height;
                    if (Label[i].name == _objectNames[labelObject])
                    {
                        output.AppendLine(Label[i].name.ToString() + div +
                            Label[i].bgColor.Name.ToString() + div +
                            Label[i].fgColor.Name.ToString() + div +
                            Label[i].width.ToString() + div +
                            Label[i].height.ToString() + div +
                            Label[i].dpi.ToString() + div +
                            Label[i].codePage.ToString() + div);
                    }
                    // text; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [default_text]; 6 fontName; 7 fontSize; 8 [fontStyle];
                    else if (Label[i].name == _objectNames[textObject])
                    {
                        output.AppendLine(Label[i].name.ToString() + div +
                            Label[i].fgColor.Name.ToString() + div +
                            Label[i].posX.ToString() + div +
                            Label[i].posY.ToString() + div +
                            Label[i].rotate.ToString() + div +
                            Label[i].content.ToString() + div +
                            Label[i].fontName.ToString() + div +
                            Label[i].fontSize.ToString() + div +
                            Label[i].fontStyle.ToString() + div);
                    }
                    // picture; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [default_file]; 6 [width]; 7 [height]; 8 [transparent];
                    else if (Label[i].name == _objectNames[pictureObject])
                    {
                        output.AppendLine(Label[i].name.ToString() + div +
                            Label[i].fgColor.Name.ToString() + div +
                            Label[i].posX.ToString() + div +
                            Label[i].posY.ToString() + div +
                            Label[i].rotate.ToString() + div +
                            Label[i].content.ToString() + div +
                            Label[i].width.ToString() + div +
                            Label[i].height.ToString() + div +
                            Convert.ToInt32(Label[i].transparent).ToString() + div);
                    }
                    // barcode; 1 [bgColor]; 2 [objectColor]; 3 posX; 4 posY; 5 [rotate]; 6 [default_data]; 7 width; 8 height; 9 bcFormat; 10 [transparent]; 11 [additional_features]
                    else if (Label[i].name == _objectNames[barcodeObject])
                    {
                        output.AppendLine(Label[i].name.ToString() + div +
                            Label[i].bgColor.Name.ToString() + div +
                            Label[i].fgColor.Name.ToString() + div +
                            Label[i].posX.ToString() + div +
                            Label[i].posY.ToString() + div +
                            Label[i].rotate.ToString() + div +
                            Label[i].content.ToString() + div +
                            Label[i].width.ToString() + div +
                            Label[i].height.ToString() + div +
                            Label[i].BCformat.ToString() + div +
                            Convert.ToInt32(Label[i].transparent).ToString() + div +
                            Label[i].feature.ToString() + div);
                    }
                    // line; 1 [objectColor]; 2 posX; 3 posY; 4 ------- ; 5 [lineWidth]; 6 endX; 7 endY; (lineLength = -1)
                    // line; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 lineLength;
                    else if (Label[i].name == _objectNames[lineObject])
                    {
                        output.Append(Label[i].name.ToString() + div +
                            Label[i].fgColor.Name.ToString() + div +
                            Label[i].posX.ToString() + div +
                            Label[i].posY.ToString() + div);
                        if (Label[i].lineLength == -1) //-V3024
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
                    // rectangle; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [transparent];
                    else if (Label[i].name == _objectNames[rectangleObject])
                    {
                        output.AppendLine(Label[i].name.ToString() + div +
                            Label[i].fgColor.Name.ToString() + div +
                            Label[i].posX.ToString() + div +
                            Label[i].posY.ToString() + div +
                            Label[i].rotate.ToString() + div +
                            Label[i].lineWidth.ToString() + div +
                            Label[i].width.ToString() + div +
                            Label[i].height.ToString() + div +
                            Convert.ToInt32(Label[i].transparent).ToString() + div);
                    }
                    // ellipse; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [transparent];
                    else if (Label[i].name == _objectNames[ellipseObject])
                    {
                        output.AppendLine(Label[i].name.ToString() + div +
                            Label[i].fgColor.Name.ToString() + div +
                            Label[i].posX.ToString() + div +
                            Label[i].posY.ToString() + div +
                            Label[i].rotate.ToString() + div +
                            Label[i].lineWidth.ToString() + div +
                            Label[i].width.ToString() + div +
                            Label[i].height.ToString() + div +
                            Convert.ToInt32(Label[i].transparent).ToString() + div);
                    }
                }
            }
            else if (SaveFileDialog1.Title == "Save label data as .CSV...")
            {
                if (checkBox_columnNames.Checked)
                {
                    for (int i = 0; i < dataGridView_labels.ColumnCount; i++) output.Append(dataGridView_labels.Columns[i].Name + div);
                    output.AppendLine();
                }
                for (int i = 0; i < dataGridView_labels.ColumnCount; i++) output.Append(dataGridView_labels.Rows[0].Cells[i].Value.ToString() + div);
                output.AppendLine();
            }
            try
            {
                File.WriteAllText(SaveFileDialog1.FileName, output.ToString(), Encoding.GetEncoding(Properties.Settings.Default.CodePage));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error writing to file " + SaveFileDialog1.FileName + ": " + ex.Message);
            }
        }

        private void comboBox_encoding_SelectedIndexChanged(object sender, EventArgs e)
        {
            template templ = new template();
            templ.name = _objectNames[labelObject];
            templ.bgColor = Label[0].bgColor;
            templ.fgColor = Label[0].fgColor;
            templ.width = Label[0].width;
            templ.height = Label[0].height;
            templ.dpi = 1;
            float.TryParse(textBox_dpi.Text, out templ.dpi);
            int cp = Properties.Settings.Default.CodePage;
            if (int.TryParse(comboBox_encoding.SelectedItem.ToString().Substring(0, comboBox_encoding.SelectedItem.ToString().IndexOf('=')), out cp)) Properties.Settings.Default.CodePage = cp;
            templ.codePage = cp;
            Label.RemoveAt(0);
            Label.Insert(0, templ);
        }

        private void textBox_dpi_Leave(object sender, EventArgs e)
        {
            template templ = new template();
            templ.name = _objectNames[labelObject];
            templ.bgColor = Label[0].bgColor;
            templ.fgColor = Label[0].fgColor;
            templ.width = Label[0].width;
            templ.height = Label[0].height;
            templ.dpi = 1;
            float.TryParse(textBox_dpi.Text, out templ.dpi);
            templ.codePage = Properties.Settings.Default.CodePage;
            if (int.TryParse(comboBox_encoding.SelectedItem.ToString().Substring(0, comboBox_encoding.SelectedItem.ToString().IndexOf('=')), out templ.codePage)) Properties.Settings.Default.CodePage = templ.codePage;
            Label.RemoveAt(0);
            Label.Insert(0, templ);
            textBox_dpi.Text = Label[0].dpi.ToString();
            units[1] = (float)(Label[0].dpi / 25.4);
            units[2] = (float)(Label[0].dpi / 2.54);
            units[3] = Label[0].dpi;
            mult = units[comboBox_units.SelectedIndex];
        }

        private void comboBox_units_SelectedIndexChanged(object sender, EventArgs e)
        {
            mult = units[comboBox_units.SelectedIndex];
            if (comboBox_units.SelectedIndex == 0) textBox_dpi.Enabled = false;
            else textBox_dpi.Enabled = true;
            showObject(listBox_objects.SelectedIndex);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex > 0 && listBox_objects.SelectedIndex < Label.Count)
            {
                if (_borderColor == Color.Black) _borderColor = Color.LightGray;
                else _borderColor = Color.Black;

                if (Label[listBox_objects.SelectedIndex].name == _objectNames[lineObject])
                    drawSelection(LabelBmp, _borderColor, currentObject.X, currentObject.Y, currentObject.Width, currentObject.Height, 0, 1);
                else
                    drawSelection(LabelBmp, _borderColor, currentObject.X, currentObject.Y, currentObject.Width, currentObject.Height, Label[listBox_objects.SelectedIndex].rotate, 1);
                pictureBox_label.Image = LabelBmp;
            }
        }

        private void drawSelection(Bitmap img, Color fgC, float posX, float posY, float width, float height, float rotateDeg, float lineWidth)
        {
            Pen p = new Pen(fgC, lineWidth);
            p.Alignment = PenAlignment.Outset;
            p.DashStyle = DashStyle.Dash;
            Rectangle rect = new Rectangle(-2, -2, (int)width + 4, (int)height + 4);
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
            // Restore the graphics state.
            g.Restore(state);
        }

        private string Evaluate(string expression)  //calculate string formula
        {
            expression = expression.Replace(',', '.');
            var loDataTable = new DataTable();
            var loDataColumn = new DataColumn("Eval", typeof(float), expression);
            loDataTable.Columns.Add(loDataColumn);
            loDataTable.Rows.Add(0);
            expression = loDataTable.Rows[0]["Eval"].ToString();
            return expression;
        }

        private void textBox_posX_Leave(object sender, EventArgs e)
        {
            textBox_posX.Text = Evaluate(textBox_posX.Text);
        }

        private void textBox_posY_Leave(object sender, EventArgs e)
        {
            textBox_posY.Text = Evaluate(textBox_posY.Text);
        }

        private void textBox_width_Leave(object sender, EventArgs e)
        {
            textBox_width.Text = Evaluate(textBox_width.Text);
        }

        private void textBox_height_Leave(object sender, EventArgs e)
        {
            textBox_height.Text = Evaluate(textBox_height.Text);
        }

        private void textBox_rotate_Leave(object sender, EventArgs e)
        {
            textBox_rotate.Text = Evaluate(textBox_rotate.Text);
        }

        private void button_clone_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < listBox_objects.Items.Count - 1 && listBox_objects.SelectedIndex > 0)
            {
                int n = listBox_objects.SelectedIndex;

                Label.Insert(listBox_objects.Items.Count - 1, Label[listBox_objects.SelectedIndex]);
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(getObjectsList());
                listBox_objects.SelectedIndex = n;
                _templateChanged = true;
                showObject(listBox_objects.SelectedIndex);
                generateLabel(-1);
            }
        }

        /*public static long EvaluateVariables(string expression, string[] variables = null, string[] values = null)  //calculate string formula
        {
            if (variables != null)
            {
                if (variables.Length != values.Length) return 0;
                for (int i = 0; i < variables.Length; i++) expression = expression.Replace(variables[i], values[i]);
            }
            var loDataTable = new DataTable();
            var loDataColumn = new DataColumn("Eval", typeof(long), expression);
            loDataTable.Columns.Add(loDataColumn);
            loDataTable.Rows.Add(0);
            return (long)(loDataTable.Rows[0]["Eval"]);
        }*/

    }
}
