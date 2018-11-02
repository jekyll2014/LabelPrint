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

        private struct Template
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

        private List<Template> Label = new List<Template>();
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
        private string printerName = null;
        private string pictureName = null;
        private Bitmap LabelBmp = new Bitmap(1, 1, PixelFormat.Format32bppPArgb);
        private List<Rectangle> currentObjectsPositions = new List<Rectangle>();
        private Color _borderColor = Color.Black;
        private string path = "";
        private Bitmap objectBmp = new Bitmap(1, 1, PixelFormat.Format32bppPArgb);

        public Form1(string[] cmdLine)
        {
            InitializeComponent();
            if (cmdLine.Length >= 1)
            {
                CmdLineOperation(cmdLine);
            }
            comboBox_object.Items.AddRange(_objectNames);

            Template init_label = new Template
            {
                name = _objectNames[labelObject],
                bgColor = Color.White,
                fgColor = Color.Black,
                width = 1,
                height = 1,
                dpi = 200,
                codePage = 65001
            };
            Label.Add(init_label);

            listBox_objects.Items.AddRange(GetObjectsList());
            listBox_objects.SelectedIndex = 0;
            comboBox_units.SelectedIndex = 0;
            textBox_dpi.Text = Label[0].dpi.ToString("F4");

            comboBox_encoding.Items.Clear();
            comboBox_encoding.Items.AddRange(GetEncodingList());

            //select codepage set in the template file
            for (int i = 0; i < comboBox_encoding.Items.Count; i++)
            {
                if (comboBox_encoding.Items[i].ToString().StartsWith(Properties.Settings.Default.CodePage.ToString()))
                {
                    comboBox_encoding.SelectedIndex = i;
                    break;
                }
            }

            TextBox_dpi_Leave(this, EventArgs.Empty);
        }

        //command line example: LabelPrint.exe /t=template.csv /l=labels.csv /c /prn="CUSTOM VKP80 II" /p=3
        private void CmdLineOperation(string[] cmdLine)
        {
            tabControl1.SelectedIndexChanged -= new EventHandler(TabControl1_SelectedIndexChanged);
            listBox_objects.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            if (cmdLine[0].StartsWith("/?") || cmdLine[0].StartsWith("/h") || cmdLine[0].StartsWith("/help"))
            {
                Console.WriteLine("/?, /h, /help - print help\r\n" +
                    "/t=file.csv - load template data from file\r\n" +
                    "/l=file.csv - load label data from file\r\n" +
                    "/c - 1st string of label file is column names (default = no)\r\n" +
                    "/prn=SystemPrinterName - output to printer (bound with \"_\" if there are spaces)\r\n" +
                    "/pic=pictureName - output to pictures\r\n" +
                    "/p=A - print all labels\r\n" +
                    "/p=xxx - print label #xxx (starts from 1)\r\n" +
                    "/p=xxx-yyy - print labels from xxx to yyy (starts from 1)");
            }
            else
            {
                string templateFile = null;
                bool columnNames = false;
                string labelFile = null;
                int _printFrom = -1;
                int _printTo = -1;
                bool printAll = false;

                for (int i = 0; i < cmdLine.Length; i++)
                {
                    cmdLine[i] = cmdLine[i].Trim();
                    if (cmdLine[i].ToLower().StartsWith("/t="))
                    {
                        templateFile = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1);
                        //check if file exists
                        if (!File.Exists(templateFile))
                        {
                            Console.WriteLine("Template file \"" + templateFile + "\" doesn't exist.");
                        }
                    }
                    else if (cmdLine[i].ToLower().StartsWith("/l="))
                    {
                        labelFile = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1);
                        //check if file exists
                        if (!File.Exists(labelFile))
                        {
                            Console.WriteLine("label file \"" + labelFile + "\" doesn't exist.");
                            labelFile = null;
                        }
                    }
                    else if (cmdLine[i].ToLower().StartsWith("/c"))
                    {
                        columnNames = true;
                    }
                    else if (cmdLine[i].ToLower().StartsWith("/prn="))
                    {
                        printerName = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1);
                        //check if printer exists
                        string[] tmp = GetPrinterList();
                        bool e = false;
                        foreach (string s in tmp)
                        {
                            if (s == printerName)
                            {
                                e = true;
                                break;
                            }
                        }
                        if (!e)
                        {
                            Console.WriteLine("label file \"" + labelFile + "\" doesn't exist.");
                            printerName = null;
                        }
                    }
                    else if (cmdLine[i].ToLower().StartsWith("/pic="))
                    {
                        pictureName = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1);
                        //check if file name correct
                    }
                    else if (cmdLine[i].ToLower().StartsWith("/p="))
                    {
                        cmdLine[i] = cmdLine[i].Substring(cmdLine[i].IndexOf('=') + 1).ToLower();
                        if (cmdLine[i] == "a")
                        {
                            printAll = true;
                            _printFrom = _printTo = 1;
                        }
                        else if (cmdLine[i].IndexOf('-') > 0)
                        {
                            if (!int.TryParse(cmdLine[i].Substring(0, cmdLine[i].IndexOf('-')), out _printFrom))
                            {
                                Console.WriteLine("Incorrect print range \"from\" parameter.");
                            }
                            if (!int.TryParse(cmdLine[i].Substring(cmdLine[i].IndexOf('-') + 1), out _printTo))
                            {
                                Console.WriteLine("Incorrect print range \"to\" parameter.");
                            }
                        }
                        else
                        {
                            if (int.TryParse(cmdLine[i], out _printFrom))
                            {
                                _printTo = _printFrom;
                            }
                            else
                            {
                                Console.WriteLine("Incorrect print range parameter.");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Unknown parameter: " + cmdLine[i]);
                    }
                }
                //check we have enough data to print
                if (templateFile != null && labelFile != null && _printFrom > 0 && _printTo > 0 && (printerName != null || pictureName != null))
                {
                    cmdLinePrint = true;
                    //import template
                    Label = LoadCsvTemplate(templateFile, Properties.Settings.Default.CodePage);

                    //import labels
                    LabelsDatabase = LoadCsvLabel(labelFile, columnNames);

                    if (printAll)
                    {
                        _printTo = LabelsDatabase.Rows.Count;
                    }
                    else if (_printTo < _printFrom)
                    {
                        int t = _printTo;
                        _printTo = _printFrom;
                        _printFrom = t;
                    }
                    else if (_printFrom > LabelsDatabase.Rows.Count)
                    {
                        Console.WriteLine("Print range oversize.\r\n");
                        printerName = null;
                        pictureName = null;
                    }
                    else if (_printTo > LabelsDatabase.Rows.Count)
                    {
                        Console.WriteLine("Print range oversize.\r\n");
                        printerName = null;
                        pictureName = null;
                    }

                    if (printerName != null) PrintLabels(_printFrom - 1, _printTo - 1, printerName);
                    if (pictureName != null) SaveLabels(_printFrom - 1, _printTo - 1, pictureName);
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

        #region Drawing

        private void FillBackground(Bitmap img, Color bgC, float width, float height)
        {
            Graphics g = Graphics.FromImage(img);
            Rectangle rect = new Rectangle(0, 0, (int)width, (int)height);
            SolidBrush b = new SolidBrush(bgC);
            g.FillRectangle(b, rect);
        }

        private Rectangle DrawText(Bitmap img, Color fgC, float posX, float posY, string text, string fontName, float fontSize, float rotateDeg = 0, FontStyle fontStyle = FontStyle.Regular)
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

        private Rectangle DrawPicture(Bitmap img, Color fgC, float posX, float posY, string fileName, float rotateDeg = 0, float width = 0, float height = 0, bool makeTransparent = true)
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

        private Rectangle DrawBarcode(Bitmap img, Color bgC, Color fgC, float posX, float posY, float width, float height, string BCdata, BarcodeFormat bcFormat, float rotateDeg = 0, string addFeature = "", bool makeTransparent = true)
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
            var barCode = new BarcodeWriter
            {
                Format = bcFormat,
                Options = new EncodingOptions
                {
                    PureBarcode = true,
                    Height = (int)height,
                    Width = (int)width
                }
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
                colorMap[0] = new ColorMap
                {
                    OldColor = Color.Black,
                    NewColor = fgC
                };
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

        private Rectangle DrawLineCoord(Bitmap img, Color fgC, float posX, float posY, float endX, float endY, float lineWidth)
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

        private Rectangle DrawLineLength(Bitmap img, Color fgC, float posX, float posY, float length, float rotateDeg, float lineWidth)
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

        private Rectangle DrawRectangle(Bitmap img, Color fgC, float posX, float posY, float width, float height, float rotateDeg, float lineWidth, bool fill)
        {
            Rectangle size = new Rectangle(0, 0, 0, 0);
            Pen p = new Pen(fgC, lineWidth)
            {
                Alignment = PenAlignment.Inset
            };
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

        private Rectangle DrawEllipse(Bitmap img, Color fgC, float posX, float posY, float width, float height, float rotateDeg, float lineWidth, bool fill)
        {
            Rectangle size = new Rectangle(0, 0, 0, 0);
            Pen p = new Pen(fgC, lineWidth)
            {
                Alignment = PenAlignment.Inset
            };
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

        private void DrawSelection(Bitmap img, Color fgC, float posX, float posY, float width, float height, float rotateDeg, float lineWidth)
        {
            Pen p = new Pen(fgC, lineWidth)
            {
                Alignment = PenAlignment.Outset,
                DashStyle = DashStyle.Dash
            };
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

        #endregion

        #region file management

        private DataTable LoadCsvLabel(string fileName, bool createColumnsNames = false)
        {
            DataTable table = new DataTable();
            //table.Clear();
            //table.Rows.Clear();
            List<string> inputStr = new List<string>();
            try
            {
                inputStr.AddRange(File.ReadAllLines(fileName, Encoding.GetEncoding(Properties.Settings.Default.CodePage)));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file:" + fileName + " : " + ex.Message);
                return table;
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
            return table;
        }

        private List<Template> LoadCsvTemplate(string fileName, int codePage)
        {
            List<Template> tmpLabel = new List<Template>();
            string[] inputStr = File.ReadAllLines(fileName, Encoding.GetEncoding(Properties.Settings.Default.CodePage));
            for (int i = 0; i < inputStr.Length; i++)
            {
                if (inputStr[i].Trim() != "")
                {
                    List<string> cells = new List<string>();
                    foreach (string str in inputStr[i].Split(Properties.Settings.Default.CSVdelimiter))
                    {
                        cells.Add(str.Trim());
                    }
                    while (cells[cells.Count - 1] == "") cells.RemoveAt(cells.Count - 1);
                    if (cells.Count >= 7)
                    {
                        Template templ = new Template();
                        if (i == 0 && cells[0] != _objectNames[labelObject])
                        {
                            MessageBox.Show("[Line " + i.ToString() + "] Incorrect or same back/foreground colors:\r\n" + inputStr[i]);
                            return tmpLabel;
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
                                    templ.fgColor = tmpLabel[0].fgColor;
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
                                else if (templ.posX >= (int)tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = (int)tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= (int)tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = (int)tmpLabel[0].height - 1;
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
                                    templ.fgColor = tmpLabel[0].fgColor;
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
                                else if (templ.posX >= (int)tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = (int)tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= (int)tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = (int)tmpLabel[0].height - 1;
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
                                /*else if (templ.width >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                    templ.width = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
                                }*/

                                float.TryParse(cells[7], out templ.height);
                                if (templ.height < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                    templ.height = 0;
                                }
                                /*else if (templ.height >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                    templ.height = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
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
                                    templ.bgColor = tmpLabel[0].bgColor;
                                }

                                if (cells[2] != "")
                                {
                                    templ.fgColor = Color.FromName(cells[2]);
                                }
                                else
                                {
                                    templ.fgColor = tmpLabel[0].fgColor;
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
                                else if (templ.posX >= (int)tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = (int)tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[4], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= (int)tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = (int)tmpLabel[0].height - 1;
                                }

                                float.TryParse(cells[5], out templ.rotate);

                                templ.content = cells[6];

                                float.TryParse(cells[7], out templ.width);
                                if (templ.width < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                    templ.width = 0;
                                }
                                /*else if (templ.width >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                    templ.width = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
                                }*/

                                float.TryParse(cells[8], out templ.height);
                                if (templ.height < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                    templ.height = 0;
                                }
                                /*else if (templ.height >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                    templ.height = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
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
                                    templ.fgColor = tmpLabel[0].fgColor;
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
                                else if (templ.posX >= (int)tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = (int)tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= (int)tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = (int)tmpLabel[0].height - 1;
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
                                    /*else if (templ.lineLength >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect line length: " + templ.lineLength.ToString());
                                        templ.lineLength = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
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
                                    /*else if (templ.width >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                        templ.width = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
                                    }*/

                                    float.TryParse(cells[7], out templ.height);
                                    if (templ.height < 0)
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = 0;
                                    }
                                    /*else if (templ.height >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                    {
                                        MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                        templ.height = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
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
                                    templ.fgColor = tmpLabel[0].fgColor;
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
                                else if (templ.posX >= (int)tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = (int)tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= (int)tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = (int)tmpLabel[0].height - 1;
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
                                /*else if (templ.width >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                    templ.width = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
                                }*/

                                float.TryParse(cells[7], out templ.height);
                                if (templ.height < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                    templ.height = 0;
                                }
                                /*else if (templ.height >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                    templ.height = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
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
                                    templ.fgColor = tmpLabel[0].fgColor;
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
                                else if (templ.posX >= (int)tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = (int)tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= (int)tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = (int)tmpLabel[0].height - 1;
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
                                /*else if (templ.width >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                    templ.width = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
                                }*/

                                float.TryParse(cells[7], out templ.height);
                                if (templ.height < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                    templ.height = 0;
                                }
                                /*else if (templ.height >= Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height))
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                    templ.height = (int)Math.Sqrt((int)tmpLabel[0].width * (int)tmpLabel[0].width + (int)tmpLabel[0].height * (int)tmpLabel[0].height);
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
                        tmpLabel.Add(templ);
                    }
                }
            }
            return tmpLabel;
        }

        private bool saveTemplateToCSV(string fileName, List<Template> _label, int codePage)
        {
            StringBuilder output = new StringBuilder();
            char div = Properties.Settings.Default.CSVdelimiter;
            for (int i = 0; i < _label.Count; i++)
            {
                // label; 1 [bgColor]; 2 [objectColor]; 3 width; 4 height;
                if (_label[i].name == _objectNames[labelObject])
                {
                    output.AppendLine(_label[i].name.ToString() + div +
                        _label[i].bgColor.Name.ToString() + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].width.ToString() + div +
                        _label[i].height.ToString() + div +
                        _label[i].dpi.ToString() + div +
                        _label[i].codePage.ToString() + div);
                }
                // text; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [default_text]; 6 fontName; 7 fontSize; 8 [fontStyle];
                else if (_label[i].name == _objectNames[textObject])
                {
                    output.AppendLine(_label[i].name.ToString() + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].posX.ToString() + div +
                        _label[i].posY.ToString() + div +
                        _label[i].rotate.ToString() + div +
                        _label[i].content.ToString() + div +
                        _label[i].fontName.ToString() + div +
                        _label[i].fontSize.ToString() + div +
                        _label[i].fontStyle.ToString() + div);
                }
                // picture; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [default_file]; 6 [width]; 7 [height]; 8 [transparent];
                else if (_label[i].name == _objectNames[pictureObject])
                {
                    output.AppendLine(_label[i].name.ToString() + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].posX.ToString() + div +
                        _label[i].posY.ToString() + div +
                        _label[i].rotate.ToString() + div +
                        _label[i].content.ToString() + div +
                        _label[i].width.ToString() + div +
                        _label[i].height.ToString() + div +
                        Convert.ToInt32(_label[i].transparent).ToString() + div);
                }
                // barcode; 1 [bgColor]; 2 [objectColor]; 3 posX; 4 posY; 5 [rotate]; 6 [default_data]; 7 width; 8 height; 9 bcFormat; 10 [transparent]; 11 [additional_features]
                else if (_label[i].name == _objectNames[barcodeObject])
                {
                    output.AppendLine(_label[i].name.ToString() + div +
                        _label[i].bgColor.Name.ToString() + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].posX.ToString() + div +
                        _label[i].posY.ToString() + div +
                        _label[i].rotate.ToString() + div +
                        _label[i].content.ToString() + div +
                        _label[i].width.ToString() + div +
                        _label[i].height.ToString() + div +
                        _label[i].BCformat.ToString() + div +
                        Convert.ToInt32(_label[i].transparent).ToString() + div +
                        _label[i].feature.ToString() + div);
                }
                // line; 1 [objectColor]; 2 posX; 3 posY; 4 ------- ; 5 [lineWidth]; 6 endX; 7 endY; (lineLength = -1)
                // line; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 lineLength;
                else if (_label[i].name == _objectNames[lineObject])
                {
                    output.Append(_label[i].name.ToString() + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].posX.ToString() + div +
                        _label[i].posY.ToString() + div);
                    if (_label[i].lineLength == -1) //-V3024
                    {
                        output.AppendLine("" + div +
                            _label[i].lineWidth.ToString() + div +
                            _label[i].width.ToString() + div +
                            _label[i].height.ToString() + div);
                    }
                    else
                    {
                        output.AppendLine(_label[i].rotate.ToString() + div +
                            _label[i].lineWidth.ToString() + div +
                            _label[i].lineLength.ToString() + div);
                    }
                }
                // rectangle; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [transparent];
                else if (_label[i].name == _objectNames[rectangleObject])
                {
                    output.AppendLine(_label[i].name.ToString() + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].posX.ToString() + div +
                        _label[i].posY.ToString() + div +
                        _label[i].rotate.ToString() + div +
                        _label[i].lineWidth.ToString() + div +
                        _label[i].width.ToString() + div +
                        _label[i].height.ToString() + div +
                        Convert.ToInt32(_label[i].transparent).ToString() + div);
                }
                // ellipse; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [transparent];
                else if (_label[i].name == _objectNames[ellipseObject])
                {
                    output.AppendLine(_label[i].name.ToString() + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].posX.ToString() + div +
                        _label[i].posY.ToString() + div +
                        _label[i].rotate.ToString() + div +
                        _label[i].lineWidth.ToString() + div +
                        _label[i].width.ToString() + div +
                        _label[i].height.ToString() + div +
                        Convert.ToInt32(_label[i].transparent).ToString() + div);
                }
            }
            bool err = true;
            try
            {
                File.WriteAllText(fileName, output.ToString(), Encoding.GetEncoding(codePage));
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error writing to file " + SaveFileDialog1.FileName + ": " + ex.Message);
                err = false;
            }
            return err;
        }

        // *BUG - saves 3 additional byte in the beginning of a file
        private bool saveTableToCSV(string fileName, DataTable dataTable, bool saveColumnNames, char csvDivider = ';', int codePage = -1)
        {
            if (codePage == -1) codePage = Encoding.UTF8.CodePage;
            StringBuilder output = new StringBuilder();
            if (saveColumnNames)
            {
                for (int i = 0; i < dataTable.Columns.Count; i++) output.Append(dataTable.Columns[i].ColumnName + csvDivider);
                output.AppendLine();
            }

            for (int j = 0; j < dataTable.Rows.Count; j++)
            {
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    output.Append(dataTable.Rows[0].ItemArray[i].ToString() + csvDivider);
                }
                output.AppendLine();
            }

            bool err = true;
            try
            {
                File.WriteAllText(fileName, output.ToString(), Encoding.GetEncoding(codePage));
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error writing to file " + SaveFileDialog1.FileName + ": " + ex.Message);
                err = false;
            }
            return err;
        }

        #endregion

        #region Result output

        private Bitmap GenerateLabel(List<Template> _label, DataTable dataTable, int lineNumber, Bitmap img)
        {
            if (lineNumber >= dataTable.Rows.Count) return null;
            if (_label[0].name != _objectNames[labelObject])
            {
                MessageBox.Show("1st object must be \"label\"");
                return null;
            }
            currentObjectsPositions.Clear();
            for (int i = 0; i < _label.Count; i++)
            {
                currentObjectsPositions.Add(new Rectangle());
                if (_label[i].name == _objectNames[labelObject])
                {
                    Color bColor = _label[i].bgColor;
                    float width = _label[i].width;
                    float height = _label[i].height;

                    img = new Bitmap((int)width, (int)height, PixelFormat.Format32bppPArgb);
                    FillBackground(img, bColor, width, height);
                    pictureBox_label.Image = img;
                }
                else if (_label[i].name == _objectNames[textObject])
                {
                    Color fColor = _label[i].fgColor;
                    float posX = _label[i].posX;
                    float posY = _label[i].posY;

                    string content = _label[i].content;
                    if (lineNumber > -1)
                        if (dataTable.Rows[lineNumber].ItemArray[i - 1].ToString() != "")
                            content = dataTable.Rows[lineNumber].ItemArray[i - 1].ToString();

                    string fontname = _label[i].fontName;
                    float fontSize = _label[i].fontSize;
                    float rotate = _label[i].rotate;
                    FontStyle fontStyle = (FontStyle)_label[i].fontStyle;

                    currentObjectsPositions[i] = DrawText(img, fColor, posX, posY, content, fontname, fontSize, rotate, fontStyle);
                }
                else if (_label[i].name == _objectNames[pictureObject])
                {
                    Color fColor = _label[i].fgColor;
                    float posX = _label[i].posX;
                    float posY = _label[i].posY;

                    string content = path + _label[i].content;
                    if (lineNumber > -1)
                        if (dataTable.Rows[lineNumber].ItemArray[i - 1].ToString() != "")
                            content = dataTable.Rows[lineNumber].ItemArray[i - 1].ToString();

                    float rotate = _label[i].rotate;
                    float width = _label[i].width;
                    float height = _label[i].height;
                    bool transparent = _label[i].transparent;

                    currentObjectsPositions[i] = DrawPicture(img, fColor, posX, posY, content, rotate, width, height, transparent);
                }
                else if (_label[i].name == _objectNames[barcodeObject])
                {
                    Color bColor = _label[i].bgColor;
                    Color fColor = _label[i].fgColor;
                    float posX = _label[i].posX;
                    float posY = _label[i].posY;
                    float width = _label[i].width;
                    float height = _label[i].height;

                    string content = _label[i].content;
                    if (lineNumber > -1)
                        if (dataTable.Rows[lineNumber].ItemArray[i - 1].ToString() != "")
                            content = dataTable.Rows[lineNumber].ItemArray[i - 1].ToString();

                    BarcodeFormat BCformat = (BarcodeFormat)_label[i].BCformat;
                    float rotate = _label[i].rotate;
                    string feature = _label[i].feature;
                    bool transparent = _label[i].transparent;

                    currentObjectsPositions[i] = DrawBarcode(img, bColor, fColor, posX, posY, width, height, content, BCformat, rotate, feature, transparent);
                }
                else if (_label[i].name == _objectNames[lineObject])
                {
                    Color fColor = _label[i].fgColor;
                    float posX = _label[i].posX;
                    float posY = _label[i].posY;
                    float rotate = _label[i].rotate;
                    float lineWidth = _label[i].lineWidth;
                    float length = _label[i].lineLength;

                    if (length == -1) //line on start/end coordinates
                    {
                        float endX = _label[i].width;
                        float endY = _label[i].height;
                        currentObjectsPositions[i] = DrawLineCoord(img, fColor, posX, posY, endX, endY, lineWidth);
                    }
                    else //line on start coordinates and length+rotate angle
                    {
                        currentObjectsPositions[i] = DrawLineLength(img, fColor, posX, posY, length, rotate, lineWidth);
                    }
                }
                else if (_label[i].name == _objectNames[rectangleObject])
                {
                    Color fColor = _label[i].fgColor;
                    float posX = _label[i].posX;
                    float posY = _label[i].posY;
                    float width = _label[i].width;
                    float height = _label[i].height;
                    float rotate = _label[i].rotate;
                    float lineWidth = _label[i].lineWidth;
                    bool fill = !_label[i].transparent;

                    currentObjectsPositions[i] = DrawRectangle(img, fColor, posX, posY, width, height, rotate, lineWidth, fill);
                }
                else if (_label[i].name == _objectNames[ellipseObject])
                {
                    Color fColor = _label[i].fgColor;
                    float posX = _label[i].posX;
                    float posY = _label[i].posY;
                    float width = _label[i].width;
                    float height = _label[i].height;
                    float rotate = _label[i].rotate;
                    float lineWidth = _label[i].lineWidth;
                    bool fill = !_label[i].transparent;

                    currentObjectsPositions[i] = DrawEllipse(img, fColor, posX, posY, width, height, rotate, lineWidth, fill);
                }
                else
                {
                    MessageBox.Show("Incorrect object: " + _label[i].name);
                    return null;
                }
            }
            //pictureBox_label.Image = (Bitmap)img.Clone();
            return img;
        }

        private void PrintLabels(int _pageFrom, int _pageTo, string prnName = "")
        {
            pagesFrom = _pageFrom;
            pagesTo = _pageTo;
            printDialog1 = new PrintDialog();
            printDocument1 = new PrintDocument();
            printDialog1.Document = printDocument1;
            printDocument1.PrintPage += new PrintPageEventHandler(PrintImagesHandler);
            if (prnName == "")
            {
                if (printDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
            }
            else
            {
                printDocument1.PrinterSettings.PrinterName = prnName;
                printDocument1.Print();
            }
        }

        private void PrintImagesHandler(object sender, PrintPageEventArgs args)
        {
            LabelBmp = GenerateLabel(Label, LabelsDatabase, pagesFrom, LabelBmp);
            pictureBox_label.Image = LabelBmp;
            if (pictureBox_label.Image == null) return;
            args.Graphics.PageUnit = GraphicsUnit.Pixel;
            Bitmap bmp = new Bitmap((int)Label[0].width, (int)Label[0].height);
            Rectangle rect = new Rectangle(0, 0, pictureBox_label.Image.Width, pictureBox_label.Image.Height);
            pictureBox_label.DrawToBitmap(bmp, rect);
            args.Graphics.DrawImage(pictureBox_label.Image, rect);
            args.HasMorePages = pagesFrom < pagesTo;
            pagesFrom++;
        }

        private void SaveLabels(int _pageFrom, int _pageTo, string filename)
        {
            for (; _pageFrom <= _pageTo; _pageFrom++)
            {
                LabelBmp = GenerateLabel(Label, LabelsDatabase, _pageFrom, LabelBmp);
                SaveLabelToFile(LabelBmp, filename + _pageFrom.ToString() + ".png");
            }
        }

        private void SaveLabelToFile(Bitmap img, string pictureName)
        {
            if (LabelBmp != null) img.Save(pictureName, ImageFormat.Png);
        }

        #endregion

        #region Utilities

        private void ClearFields()
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

        private Template CollectObjectFromGUI(int n)
        {
            Template templ = new Template();
            if (n < 0) return templ;

            if (n == listBox_objects.Items.Count - 1)
            {
                templ.name = comboBox_object.SelectedItem.ToString();
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
            else
            {
                templ = Label[n];
                if (templ.name == _objectNames[labelObject])
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
                    textBox_fontSize.Text = Evaluate(textBox_fontSize.Text).ToString();
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

                    textBox_fontSize.Text = Evaluate(textBox_fontSize.Text).ToString();
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

                        textBox_content.Text = Evaluate(textBox_content.Text).ToString();
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

                    textBox_fontSize.Text = Evaluate(textBox_fontSize.Text).ToString();
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

                    textBox_fontSize.Text = Evaluate(textBox_fontSize.Text).ToString();
                    float.TryParse(textBox_fontSize.Text, out f);
                    templ.lineWidth = f * mult;

                    templ.transparent = !checkBox_fill.Checked;
                }
            }
            return templ;
        }

        private void ShowObjectInGUI(int n)
        {
            if (n < 0) return;
            ClearFields();

            string str = "";
            //new object
            if (n >= Label.Count)
            {
                comboBox_object.Enabled = true;
            }
            else
            {
                comboBox_object.SelectedItem = Label[n].name;
                // label; [bgColor]; [objectColor]; width; height;
                if (Label[n].name == _objectNames[labelObject])
                {
                    label_backgroundColor.Text = "Background color";
                    comboBox_backgroundColor.Enabled = true;
                    comboBox_backgroundColor.Items.Clear();

                    comboBox_backgroundColor.Items.AddRange(GetColorList());
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
                    comboBox_objectColor.Items.AddRange(GetColorList());
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
                    comboBox_objectColor.Items.AddRange(GetColorList());
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
                    comboBox_fontName.Items.AddRange(GetFontList());
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
                    comboBox_objectColor.Items.AddRange(GetColorList());
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
                    comboBox_backgroundColor.Items.AddRange(GetColorList());
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
                    comboBox_objectColor.Items.AddRange(GetColorList());
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
                    comboBox_fontName.Items.AddRange(GetBarcodeList());
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
                    comboBox_objectColor.Items.AddRange(GetColorList());
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
                    label_width.Text = "endX";
                    textBox_width.Text = (Label[n].width / mult).ToString("F4");

                    textBox_height.Enabled = true;
                    label_height.Text = "endY";
                    textBox_height.Text = (Label[n].height / mult).ToString("F4");
                    textBox_rotate.Enabled = true;
                    label_rotate.Text = "Rotate";
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
                    comboBox_objectColor.Items.AddRange(GetColorList());
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
                    comboBox_objectColor.Items.AddRange(GetColorList());
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
            }
        }

        private string[] GetObjectsList()
        {
            List<string> objectList = new List<string>();
            int i = 0;
            foreach (Template t in Label)
            {
                objectList.Add(i.ToString() + " " + t.name);
                i++;
            }
            objectList.Add("");
            return objectList.ToArray();
        }

        private string[] GetColorList()
        {
            List<string> colorList = new List<string>();
            foreach (Color c in new ColorConverter().GetStandardValues())
            {
                colorList.Add(c.Name);
            }
            return colorList.ToArray();
        }

        private string[] GetFontList()
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

        private string[] GetBarcodeList()
        {
            List<string> barcodeList = new List<string>();
            foreach (BarcodeFormat b in BarCodeTypes)
            {
                barcodeList.Add(((int)b).ToString() + "=" + b.ToString());
            }
            return barcodeList.ToArray();
        }

        private string[] GetEncodingList()
        {
            List<string> encodeList = new List<string>();

            EncodingInfo[] codePages = Encoding.GetEncodings();
            foreach (EncodingInfo codePage in codePages)
            {
                encodeList.Add(codePage.CodePage.ToString() + "=" + "[" + codePage.Name + "]" + codePage.DisplayName);
            }
            return encodeList.ToArray();
        }

        private string[] GetPrinterList()
        {
            List<string> printerList = new List<string>();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                printerList.Add(printer);
            }
            return printerList.ToArray();
        }

        private void SetRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
            dgv.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private float Evaluate(string expression)  //calculate string formula
        {
            expression = expression.Replace(',', '.');
            var loDataTable = new DataTable();
            var loDataColumn = new DataColumn("Eval", typeof(float), expression);
            loDataTable.Columns.Add(loDataColumn);
            loDataTable.Rows.Add(0);
            float e = (float)loDataTable.Rows[0]["Eval"];
            return e;
        }

        #endregion

        #region GUI management

        private void Button_importLabels_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Title = "Open labels .CSV database";
            openFileDialog1.DefaultExt = "csv";
            openFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
            openFileDialog1.ShowDialog();
        }

        private void Button_importTemplate_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Title = "Open template .CSV file";
            openFileDialog1.DefaultExt = "csv";
            openFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
            openFileDialog1.ShowDialog();
        }

        private void OpenFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (openFileDialog1.Title == "Open template .CSV file")
            {
                objectsNum = 0;
                dataGridView_labels.DataSource = null;
                LabelsDatabase.Clear();
                LabelsDatabase.Columns.Clear();
                LabelsDatabase.Rows.Clear();
                textBox_labelsName.Clear();
                Label.Clear();

                Label = LoadCsvTemplate(openFileDialog1.FileName, Properties.Settings.Default.CodePage);

                TextBox_dpi_Leave(this, EventArgs.Empty);
                button_importLabels.Enabled = true;
                LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
                pictureBox_label.Image = LabelBmp;

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
                foreach (DataGridViewColumn column in dataGridView_labels.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
                button_printCurrent.Enabled = true;
                button_printAll.Enabled = false;
                button_printRange.Enabled = false;
                textBox_rangeFrom.Text = "0";
                textBox_rangeTo.Text = "0";
                SetRowNumber(dataGridView_labels);
                TabControl1_SelectedIndexChanged(this, EventArgs.Empty);
            }
            else if (openFileDialog1.Title == "Open labels .CSV database")
            {
                dataGridView_labels.DataSource = null;

                LabelsDatabase = LoadCsvLabel(openFileDialog1.FileName, checkBox_columnNames.Checked);

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
                    textBox_rangeTo.Text = (LabelsDatabase.Rows.Count).ToString();
                    SetRowNumber(dataGridView_labels);
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
                LabelBmp = GenerateLabel(Label, LabelsDatabase, dataGridView_labels.CurrentCell.RowIndex, LabelBmp);
                pictureBox_label.Image = LabelBmp;
                textBox_labelsName.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\') + 1);
            }
        }

        private void Button_saveTemplate_Click(object sender, EventArgs e)
        {
            SaveFileDialog1.Title = "Save template as .CSV...";
            SaveFileDialog1.DefaultExt = "csv";
            SaveFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
            if (textBox_templateName.Text == "") SaveFileDialog1.FileName = "template_" + DateTime.Today.ToShortDateString().Replace("/", "_") + ".csv";
            else SaveFileDialog1.FileName = textBox_templateName.Text;
            SaveFileDialog1.ShowDialog();
        }

        private void Button_saveLabel_Click(object sender, EventArgs e)
        {
            SaveFileDialog1.Title = "Save label data as .CSV...";
            SaveFileDialog1.DefaultExt = "csv";
            SaveFileDialog1.Filter = "CSV files|*.csv|All files|*.*";
            SaveFileDialog1.FileName = "label_" + textBox_templateName.Text;
            SaveFileDialog1.ShowDialog();
        }

        private void SaveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StringBuilder output = new StringBuilder();
            char div = Properties.Settings.Default.CSVdelimiter;
            if (SaveFileDialog1.Title == "Save template as .CSV...")
            {
                if (!saveTemplateToCSV(SaveFileDialog1.FileName, Label, Properties.Settings.Default.CodePage)) MessageBox.Show("Error writing to file " + SaveFileDialog1.FileName);
            }
            else if (SaveFileDialog1.Title == "Save label data as .CSV...")
            {
                if (!saveTableToCSV(SaveFileDialog1.FileName, LabelsDatabase, checkBox_columnNames.Checked, Properties.Settings.Default.CSVdelimiter, Properties.Settings.Default.CodePage)) MessageBox.Show("Error writing to file " + SaveFileDialog1.FileName);
            }
        }

        private void CheckBox_toFile_CheckedChanged(object sender, EventArgs e)
        {
            textBox_saveFileName.Enabled = checkBox_toFile.Checked;
        }

        private void CheckBox_scale_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_scale.Checked)
            {
                //pictureBox_label.Dock = DockStyle.None;
                pictureBox_label.SizeMode = PictureBoxSizeMode.Normal;
                //pictureBox_label.Width = (int)Label[0].width;
                //pictureBox_label.Height = (int)Label[0].height;
            }
            else
            {
                //pictureBox_label.Dock = DockStyle.Fill;
                pictureBox_label.SizeMode = PictureBoxSizeMode.Zoom;
            }
        }

        private void DataGridView_labels_SelectionChanged(object sender, EventArgs e)
        {
            if (!cmdLinePrint)
            {
                LabelBmp = GenerateLabel(Label, LabelsDatabase, dataGridView_labels.CurrentCell.RowIndex, LabelBmp);
                pictureBox_label.Image = LabelBmp;
            }
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex != 0)
            {
                timer1.Enabled = true;
            }
            if (tabControl1.SelectedIndex == 0)
            {
                timer1.Enabled = false;
                if (_templateChanged)
                {
                    dataGridView_labels.SelectionChanged -= new EventHandler(DataGridView_labels_SelectionChanged);
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
                    textBox_rangeTo.Text = LabelsDatabase.Rows.Count.ToString();
                    SetRowNumber(dataGridView_labels);
                    dataGridView_labels.CurrentCell = dataGridView_labels.Rows[0].Cells[0];
                    dataGridView_labels.Rows[0].Selected = true;
                    //GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
                    //pictureBox_label.Image = LabelBmp;
                    dataGridView_labels.SelectionChanged += new EventHandler(DataGridView_labels_SelectionChanged);
                    _templateChanged = false;
                }
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(GetObjectsList());
                listBox_objects.SelectedIndex = 0;
                listBox_objects.SelectedIndex = 0;
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                listBox_objectsMulti.Items.Clear();
                listBox_objectsMulti.Items.AddRange(GetObjectsList());
                listBox_objectsMulti.SelectedIndex = 0;
            }
            LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
            pictureBox_label.Image = LabelBmp;
        }

        private void ListBox_objects_SelectedIndexChanged(object sender, EventArgs e)
        {
            ShowObjectInGUI(listBox_objects.SelectedIndex);
            LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
            pictureBox_label.Image = LabelBmp;
        }

        private void Button_apply_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < 0) return;
            listBox_objects.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            int n = listBox_objects.SelectedIndex;
            Template templ = CollectObjectFromGUI(n);

            if (n >= Label.Count)
            {
                if (comboBox_object.SelectedIndex > 0) Label.Add(templ);
            }
            else
            {
                Label[n] = templ;
            }
            LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
            pictureBox_label.Image = LabelBmp;
            listBox_objects.Items.Clear();
            listBox_objects.Items.AddRange(GetObjectsList());
            listBox_objects.SelectedIndex = n;
            _templateChanged = true;
            ShowObjectInGUI(listBox_objects.SelectedIndex);
            listBox_objects.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_delete_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < listBox_objects.Items.Count - 1 && listBox_objects.SelectedIndex > 0)
            {
                int n = listBox_objects.SelectedIndex;
                Label.RemoveAt(n);
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(GetObjectsList());
                listBox_objects.SelectedIndex = n;
                _templateChanged = true;
                ShowObjectInGUI(listBox_objects.SelectedIndex);
                LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
                pictureBox_label.Image = LabelBmp;
            }
        }

        private void Button_up_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < listBox_objects.Items.Count - 1 && listBox_objects.SelectedIndex > 1)
            {
                int n = listBox_objects.SelectedIndex;
                Template templ = Label[n];
                Label.RemoveAt(n);
                Label.Insert(n - 1, templ);
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(GetObjectsList());
                listBox_objects.SelectedIndex = n - 1;
                _templateChanged = true;
                ShowObjectInGUI(listBox_objects.SelectedIndex);
                GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
                pictureBox_label.Image = LabelBmp;
            }
        }

        private void Button_down_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < listBox_objects.Items.Count - 2 && listBox_objects.SelectedIndex > 0)
            {
                int n = listBox_objects.SelectedIndex;
                Template templ = Label[n];
                Label.RemoveAt(n);
                Label.Insert(n + 1, templ);
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(GetObjectsList());
                listBox_objects.SelectedIndex = n + 1;
                _templateChanged = true;
                ShowObjectInGUI(listBox_objects.SelectedIndex);
                LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
                pictureBox_label.Image = LabelBmp;
            }
        }

        private void ComboBox_encoding_SelectedIndexChanged(object sender, EventArgs e)
        {
            Template templ = new Template
            {
                name = _objectNames[labelObject],
                bgColor = Label[0].bgColor,
                fgColor = Label[0].fgColor,
                width = Label[0].width,
                height = Label[0].height,
                dpi = 1
            };
            float.TryParse(textBox_dpi.Text, out templ.dpi);
            int cp = Properties.Settings.Default.CodePage;
            if (int.TryParse(comboBox_encoding.SelectedItem.ToString().Substring(0, comboBox_encoding.SelectedItem.ToString().IndexOf('=')), out cp)) Properties.Settings.Default.CodePage = cp;
            templ.codePage = cp;
            Label.RemoveAt(0);
            Label.Insert(0, templ);
        }

        private void TextBox_dpi_Leave(object sender, EventArgs e)
        {
            Template templ = new Template
            {
                name = _objectNames[labelObject],
                bgColor = Label[0].bgColor,
                fgColor = Label[0].fgColor,
                width = Label[0].width,
                height = Label[0].height,
                dpi = 1
            };
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

        private void ComboBox_units_SelectedIndexChanged(object sender, EventArgs e)
        {
            mult = units[comboBox_units.SelectedIndex];
            if (comboBox_units.SelectedIndex == 0) textBox_dpi.Enabled = false;
            else textBox_dpi.Enabled = true;
            ShowObjectInGUI(listBox_objects.SelectedIndex);
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            if (_borderColor == Color.Black) _borderColor = Color.LightGray;
            else _borderColor = Color.Black;

            var list = listBox_objects.SelectedIndices;
            if (tabControl1.SelectedIndex == 2) list = listBox_objectsMulti.SelectedIndices;

            foreach (int n in list)
            {
                if (n > 0 && n < Label.Count)
                {

                    if (Label[n].name == _objectNames[lineObject])
                        DrawSelection(LabelBmp, _borderColor, currentObjectsPositions[n].X, currentObjectsPositions[n].Y, currentObjectsPositions[n].Width, currentObjectsPositions[n].Height, 0, 1);
                    else
                        try
                        {
                            DrawSelection(LabelBmp, _borderColor, currentObjectsPositions[n].X, currentObjectsPositions[n].Y, currentObjectsPositions[n].Width, currentObjectsPositions[n].Height, Label[n].rotate, 1);
                        }
                        catch { }
                }
            }
            pictureBox_label.Image = LabelBmp;
        }

        private void TextBox_posX_Leave(object sender, EventArgs e)
        {
            textBox_posX.Text = Evaluate(textBox_posX.Text).ToString();
        }

        private void TextBox_posY_Leave(object sender, EventArgs e)
        {
            textBox_posY.Text = Evaluate(textBox_posY.Text).ToString();
        }

        private void TextBox_width_Leave(object sender, EventArgs e)
        {
            textBox_width.Text = Evaluate(textBox_width.Text).ToString();
        }

        private void TextBox_height_Leave(object sender, EventArgs e)
        {
            textBox_height.Text = Evaluate(textBox_height.Text).ToString();
        }

        private void TextBox_rotate_Leave(object sender, EventArgs e)
        {
            textBox_rotate.Text = Evaluate(textBox_rotate.Text).ToString();
        }

        private void Button_clone_Click(object sender, EventArgs e)
        {
            if (listBox_objects.SelectedIndex < listBox_objects.Items.Count - 1 && listBox_objects.SelectedIndex > 0)
            {
                int n = listBox_objects.SelectedIndex;

                Label.Insert(listBox_objects.Items.Count - 1, Label[listBox_objects.SelectedIndex]);
                listBox_objects.Items.Clear();
                listBox_objects.Items.AddRange(GetObjectsList());
                listBox_objects.SelectedIndex = n;
                _templateChanged = true;
                ShowObjectInGUI(listBox_objects.SelectedIndex);
                LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
                pictureBox_label.Image = LabelBmp;
            }
        }

        private void Button_moveUp_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1)
                {
                    Template templ = Label[n];
                    float y = 0;
                    float.TryParse(textBox_move.Text, out y);
                    templ.posY -= y;
                    Label.RemoveAt(n);
                    Label.Insert(n, templ);
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.Items.Clear();
            listBox_objectsMulti.Items.AddRange(GetObjectsList());
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_moveLeft_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1)
                {
                    Template templ = Label[n];
                    float x = 0;
                    float.TryParse(textBox_move.Text, out x);
                    templ.posX -= x;
                    Label.RemoveAt(n);
                    Label.Insert(n, templ);
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.Items.Clear();
            listBox_objectsMulti.Items.AddRange(GetObjectsList());
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_moveDown_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1)
                {
                    Template templ = Label[n];
                    float y = 0;
                    float.TryParse(textBox_move.Text, out y);
                    templ.posY += y;
                    Label.RemoveAt(n);
                    Label.Insert(n, templ);
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.Items.Clear();
            listBox_objectsMulti.Items.AddRange(GetObjectsList());
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_moveRight_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1)
                {
                    Template templ = Label[n];
                    float x = 0;
                    float.TryParse(textBox_move.Text, out x);
                    templ.posX += x;
                    Label.RemoveAt(n);
                    Label.Insert(n, templ);
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.Items.Clear();
            listBox_objectsMulti.Items.AddRange(GetObjectsList());
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_deleteGroup_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1)
                {
                    Label.RemoveAt(n);
                    _templateChanged = true;
                }
            }
            listBox_objectsMulti.Items.Clear();
            listBox_objectsMulti.Items.AddRange(GetObjectsList());
            listBox_objectsMulti.SelectedIndex = 0;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
            pictureBox_label.Image = LabelBmp;
        }

        private void TextBox_move_Leave(object sender, EventArgs e)
        {
            float n = 0;
            float.TryParse(textBox_move.Text, out n);
            textBox_move.Text = n.ToString();
        }

        private void textBox_rangeFrom_Leave(object sender, EventArgs e)
        {
            int n = 1;
            int.TryParse(textBox_rangeFrom.Text, out n);
            if (n >= 1) textBox_rangeFrom.Text = n.ToString();
        }

        private void textBox_rangeTo_Leave(object sender, EventArgs e)
        {
            int n = LabelsDatabase.Rows.Count;
            int.TryParse(textBox_rangeTo.Text, out n);
            if (n <= LabelsDatabase.Rows.Count) textBox_rangeTo.Text = n.ToString();
        }

        private void ListBox_objectsMulti_SelectedIndexChanged(object sender, EventArgs e)
        {
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
            pictureBox_label.Image = LabelBmp;
        }

        private void Button_printCurrent_Click(object sender, EventArgs e)
        {
            int _pagesFrom = dataGridView_labels.CurrentRow.Index;
            int _pagesTo = _pagesFrom;

            if (!checkBox_toFile.Checked)
            {
                PrintLabels(_pagesFrom, _pagesTo);
            }
            else
            {
                SaveLabels(_pagesFrom, _pagesTo, textBox_saveFileName.Text);
            }
        }

        private void Button_printAll_Click(object sender, EventArgs e)
        {
            int _pagesFrom = 0;
            int _pagesTo = LabelsDatabase.Rows.Count - 1;

            if (!checkBox_toFile.Checked)
            {
                PrintLabels(_pagesFrom, _pagesTo);
            }
            else
            {
                SaveLabels(_pagesFrom, _pagesTo, textBox_saveFileName.Text);
            }
        }

        private void Button_printRange_Click(object sender, EventArgs e)
        {
            int _pagesFrom = 0;
            int _pagesTo = LabelsDatabase.Rows.Count;

            int.TryParse(textBox_rangeFrom.Text, out _pagesFrom);
            int.TryParse(textBox_rangeTo.Text, out _pagesTo);
            if (_pagesFrom < 1) _pagesFrom = 1;
            if (_pagesTo > LabelsDatabase.Rows.Count) _pagesTo = LabelsDatabase.Rows.Count;

            if (!checkBox_toFile.Checked)
            {
                PrintLabels(_pagesFrom - 1, _pagesTo - 1);
            }
            else
            {
                SaveLabels(_pagesFrom - 1, _pagesTo - 1, textBox_saveFileName.Text);
            }
        }

        #endregion

        #region Drag'n'Drop

        private void PictureBox_label_MouseMove(object sender, MouseEventArgs e)
        {
            if (pictureBox_label.Image == null) return;
            if (checkBox_scale.Checked)
            {

                textBox_mX.Text = e.X.ToString();
                textBox_mY.Text = e.Y.ToString();
            }
            else
            {
                float scaleW = (float)pictureBox_label.Width / LabelBmp.Width;
                float scaleH = (float)pictureBox_label.Height / LabelBmp.Height;
                if (scaleW > scaleH)
                {
                    textBox_mY.Text = (e.Y / scaleH).ToString();
                    float corr = (pictureBox_label.Width - (LabelBmp.Width * scaleH)) / 2;
                    textBox_mX.Text = ((e.X - corr) / scaleH).ToString();
                }
                else
                {
                    textBox_mX.Text = (e.X / scaleW).ToString();
                    float corr = (pictureBox_label.Height - LabelBmp.Height * scaleW) / 2;
                    textBox_mY.Text = ((e.Y - corr) / scaleW).ToString();
                }
            }
        }

        /*int xPosOrig, yPosOrig;
        int xPosCurrent, yPosCurrent;
        private void pictureBox_ball_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                xPosOrig = e.X;
                yPosOrig = e.Y;
            }
        }*/

        /*private void pictureBox_ball_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox p = sender as PictureBox;
                if (p != null)
                {
                    xPosCurrent = pictureBox_field.Left + pictureBox_field.Width / 2 - p.Width / 2;
                    yPosCurrent = pictureBox_field.Top + pictureBox_field.Height / 2 - p.Height / 2;
                    p.Left = xPosCurrent;
                    p.Top = yPosCurrent;
                }
                //send stop signal to cart
                if (serialPort1.IsOpen)
                {
                    try
                    {
                        serialPort1.Write("keep 0 0\r\n");
                        serialPort1.Write("s\r\n");
                    }
                    catch { }
                }
                textBox_jX.Text = "0";
                textBox_jY.Text = "0";
                L_sent = 0;
                R_sent = 0;
            }
        }*/

        /*private void pictureBox_ball_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox p = sender as PictureBox;
                if (p != null)
                {
                    xPosCurrent += e.X - xPosOrig;
                    if (xPosCurrent < pictureBox_field.Left) xPosCurrent = pictureBox_field.Left;
                    if (xPosCurrent > pictureBox_field.Left + pictureBox_field.Width - p.Width) xPosCurrent = pictureBox_field.Left + pictureBox_field.Width - p.Width;
                    p.Left = xPosCurrent;

                    yPosCurrent += e.Y - yPosOrig;
                    if (yPosCurrent < pictureBox_field.Top) yPosCurrent = pictureBox_field.Top;
                    if (yPosCurrent > pictureBox_field.Top + pictureBox_field.Height - p.Height) yPosCurrent = pictureBox_field.Top + pictureBox_field.Height - p.Height;
                    p.Top = yPosCurrent;

                    float jX = (xPosCurrent - xPosOrig - (pictureBox_field.Left + (float)pictureBox_field.Width / 2 - (float)pictureBox_ball.Width / 2) + (float)pictureBox_ball.Width / 2) / ((float)pictureBox_field.Width / 2);
                    float jY = -((yPosCurrent - yPosOrig - (pictureBox_field.Top + (float)pictureBox_field.Height / 2 - (float)pictureBox_ball.Height / 2) + (float)pictureBox_ball.Height / 2) / ((float)pictureBox_field.Height / 2));

                    //calc speed
                    float l = 0;
                    float r = 0;
                    float spd = jY * max_speed * 10;

                    if (jX > 0)
                    {
                        l = (int)(spd);
                        r = (int)(spd * (1 - jX));
                    }
                    else
                    {
                        r = (int)spd;
                        l = (int)(spd * (1 + jX));
                    }
                    l = l / 10;
                    r = r / 10;
                    //if difference >0.1 send speed to cart
                    if (serialPort1.IsOpen)
                    {
                        if (Math.Abs(L_sent - l) >= 0.1 || Math.Abs(R_sent - r) >= 0.1)
                        {
                            try
                            {
                                serialPort1.Write("keep " + l.ToString("F2") + " " + r.ToString("F2") + "\r\n");
                            }
                            catch { }
                        }
                    }
                    //redraw L/R speed indicator on the picture
                    textBox_jX.Text = L_sent.ToString("F2");
                    textBox_jY.Text = R_sent.ToString("F2");
                    L_sent = l;
                    R_sent = r;
                }
            }
        }*/

        #endregion
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

/*private bool IsInPolygon(Point[] poly, Point pnt)
{
    int i, j;
    int nvert = poly.Length;
    bool c = false;
    for (i = 0, j = nvert - 1; i < nvert; j = i++)
    {
        if (((poly[i].Y > pnt.Y) != (poly[j].Y > pnt.Y)) &&
         (pnt.X < (poly[j].X - poly[i].X) * (pnt.Y - poly[i].Y) / (poly[j].Y - poly[i].Y) + poly[i].X))
            c = !c;
    }
    return c;
}*/

/*private PointF[] rotatePolygon(PointF zeroPoint, PointF[] poly, float angle)
{
    PointF[] p = new PointF[poly.Length];
    for (int i = 0; i < poly.Length; i++)
    {
        double xn = 0, yn = 0;
        rotateLine(zeroPoint.X, zeroPoint.Y, poly[i].X, poly[i].Y, angle, out xn, out yn);
        p[i].X = (float)xn;
        p[i].X = (float)yn;
    }
    return p;
}*/

/*private void rotateLine(double xc, double yc, double x, double y, double a, out double xn, out double yn)
{
    xn = (x - xc) * Math.Cos(a) - (y - yc) * Math.Sin(a) + xc;
    yn = (x - xc) * Math.Sin(a) + (y - yc) * Math.Cos(a) + yc;
}*/
