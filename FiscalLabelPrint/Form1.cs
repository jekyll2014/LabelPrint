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
        private const int line_coord_Object = 4;
        private const int line_length_Object = 5;
        private const int rectangleObject = 6;
        private const int ellipseObject = 7;

        private enum LabelObject
        {
            label,
            text,
            picture,
            barcode,
            line_coord,
            line_length,
            rectangle,
            ellipse,
        }

        private struct Template
        {
            public LabelObject objectType;
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
            public BarcodeFormat barCodeFormat;
            public float fontSize;
            public FontStyle fontStyle;
            public string fontName;
            public string addFeature;
            public float lineLength;
            public float lineWidth;
        }

        private List<Template> Label = new List<Template>();
        private DataTable LabelsDatabase = new DataTable();
        private List<string> bcFeatures = new List<string>();

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
        private float mouseX = 0, mouseY = 0;

        public Form1(string[] cmdLine)
        {
            InitializeComponent();
            if (cmdLine.Length >= 1)
            {
                CmdLineOperation(cmdLine);
            }
            comboBox_object.Items.AddRange(Enum.GetNames(typeof(LabelObject)));

            Template init_label = new Template
            {
                objectType = LabelObject.label,
                bgColor = Color.White,
                fgColor = Color.Black,
                width = 100,
                height = 100,
                dpi = 300,
                codePage = Properties.Settings.Default.CodePage
            };
            Label.Add(init_label);

            bcFeatures.AddRange(Enum.GetNames(typeof(EncodeHintType)));
            listBox_objects.Items.AddRange(GetObjectsList());
            listBox_objects.SelectedIndex = 0;
            comboBox_units.SelectedIndex = 0;
            textBox_dpi.Text = Label[0].dpi.ToString("F4");

            comboBox_encoding.Items.Clear();
            comboBox_encoding.Items.AddRange(GetEncodingList());

            //select codepage set in the template file
            for (int i = 0; i < comboBox_encoding.Items.Count; i++)
            {
                if (comboBox_encoding.Items[i].ToString().StartsWith(Label[0].codePage.ToString()))
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
            if (cmdLine[0].StartsWith("/?") || cmdLine[0].StartsWith("/help"))
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
                    path = templateFile.Substring(0, templateFile.LastIndexOf('\\') + 1);
                    Label = LoadCsvTemplate(templateFile, Label[0].codePage);
                    if (Label.Count <= 1)
                    {
                        Console.WriteLine("Incorrect template file.\r\n");
                        Exit();
                    }
                    //import labels
                    path = templateFile.Substring(0, templateFile.LastIndexOf('\\') + 1);
                    LabelsDatabase = LoadCsvLabel(labelFile, Label[0].codePage, columnNames);
                    if (LabelsDatabase.Rows.Count < 1)
                    {
                        Console.WriteLine("Incorrect label file.\r\n");
                        Exit();
                    }

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
            Exit();
        }

        private void Exit()
        {
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
            if (addFeature != "" && addFeature.Contains("="))
            {
                string hint = addFeature.Substring(0, addFeature.IndexOf("=")).Trim();
                if (hint.Length > 0)
                {
                    Enum.TryParse<EncodeHintType>(hint, out EncodeHintType _feature);
                    string value = addFeature.Substring(addFeature.IndexOf("=") + 1).Trim();
                    switch (_feature)
                    {
                        /* Specifies the required number of layers for an Aztec code. A negative number (-1, -2, -3, -4) specifies a compact Aztec
                         * code 0 indicates to use the minimum number of layers (the default) A positive number (1, 2, .. 32) specifies a normal
                         * (non-compact) Aztec code type: System.Int32, or System.String representation of the integer value
                         */
                        case EncodeHintType.AZTEC_LAYERS: //int [-4, 32]
                            {
                                _feature = EncodeHintType.AZTEC_LAYERS;
                                if (int.TryParse(value, out int param))
                                {
                                    if (param < -4) param = -4;
                                    else if (param > 32) param = 32;
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Specifies what character encoding to use where applicable. type: System.String
                         */
                        case EncodeHintType.CHARACTER_SET: //string
                            {
                                _feature = EncodeHintType.CHARACTER_SET;
                                barCode.Options.Hints.Add(_feature, value);
                            }
                            break;

                        /* if true, don't switch to codeset C for numbers
                         */
                        case EncodeHintType.CODE128_FORCE_CODESET_B: //bool
                            {
                                _feature = EncodeHintType.CODE128_FORCE_CODESET_B;
                                if (bool.TryParse(value, out bool param))
                                {
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Specifies the default encodation for Data Matrix (type ZXing.Datamatrix.Encoder.Encodation) Make sure that the content fits into
                         * the encodation value, otherwise there will be an exception thrown. standard value: Encodation.ASCII
                         */
                        case EncodeHintType.DATA_MATRIX_DEFAULT_ENCODATION: //int [0,1,2,3,4,5] (ASCII,C40,TEXT,X12,EDIFACT,BASE256)
                            {
                                _feature = EncodeHintType.DATA_MATRIX_DEFAULT_ENCODATION;
                                if (int.TryParse(value, out int param))
                                {
                                    if (param < 0) param = 0;
                                    else if (param > 5) param = 5;
                                    if (param == (int)ZXing.Datamatrix.Encoder.Encodation.ASCII)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.Datamatrix.Encoder.Encodation.ASCII);
                                    }
                                    else if (param == (int)ZXing.Datamatrix.Encoder.Encodation.BASE256)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.Datamatrix.Encoder.Encodation.BASE256);
                                    }
                                    else if (param == (int)ZXing.Datamatrix.Encoder.Encodation.C40)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.Datamatrix.Encoder.Encodation.C40);
                                    }
                                    else if (param == (int)ZXing.Datamatrix.Encoder.Encodation.EDIFACT)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.Datamatrix.Encoder.Encodation.EDIFACT);
                                    }
                                    else if (param == (int)ZXing.Datamatrix.Encoder.Encodation.TEXT)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.Datamatrix.Encoder.Encodation.TEXT);
                                    }
                                    else if (param == (int)ZXing.Datamatrix.Encoder.Encodation.X12)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.Datamatrix.Encoder.Encodation.X12);
                                    }
                                }
                            }
                            break;

                        /* Specifies the matrix shape for Data Matrix (type ZXing.Datamatrix.Encoder.SymbolShapeHint)
                         */
                        case EncodeHintType.DATA_MATRIX_SHAPE: //string [FORCE_NONE, FORCE_RECTANGLE, FORCE_SQUARE]
                            {
                                _feature = EncodeHintType.DATA_MATRIX_SHAPE;
                                if (Enum.TryParse<ZXing.Datamatrix.Encoder.SymbolShapeHint>(value, out var param))
                                {
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Don't append ECI segment. That is against the specification of QR Code but some readers have problems if the charset is switched
                         * from ISO-8859-1 (default) to UTF-8 with the necessary ECI segment. If you set the property to true you can use UTF-8 encoding and
                         * the ECI segment is omitted. type: System.Boolean
                         */
                        case EncodeHintType.DISABLE_ECI: //bool
                            {
                                _feature = EncodeHintType.DISABLE_ECI;
                                if (bool.TryParse(value, out bool param))
                                {
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Specifies what degree of error correction to use, for example in QR Codes. Type depends on the encoder. For example for QR codes
                         * it's type ZXing.QrCode.Internal.ErrorCorrectionLevel For Aztec it is of type System.Int32, representing the minimal percentage of
                         * error correction words. In all cases, it can also be a System.String representation of the desired value as well. Note: an Aztec
                         * symbol should have a minimum of 25% EC words. For PDF417 it is of type ZXing.PDF417.Internal.PDF417ErrorCorrectionLevel or
                         * System.Int32 (between 0 and 8),
                         */
                        case EncodeHintType.ERROR_CORRECTION: //int PDF417[0, 8]; AZTEC [25-100]%; QRCODE [L,M,Q,H] (7%,15%,25%,30%)
                            {
                                _feature = EncodeHintType.ERROR_CORRECTION;
                                if (barCode.Format == BarcodeFormat.AZTEC)
                                {
                                    if (int.TryParse(value, out int param))
                                    {
                                        if (param < 25) param = 25;
                                        else if (param > 100) param = 100;
                                        barCode.Options.Hints.Add(_feature, param);
                                    }
                                }
                                else if (barCode.Format == BarcodeFormat.PDF_417)
                                {
                                    if (int.TryParse(value, out int param))
                                    {
                                        if (param < 0) param = 0;
                                        else if (param > 8) param = 8;
                                        barCode.Options.Hints.Add(_feature, param);
                                    }
                                }
                                else if (barCode.Format == BarcodeFormat.QR_CODE)
                                {
                                    if (value == ZXing.QrCode.Internal.ErrorCorrectionLevel.L.Name)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.QrCode.Internal.ErrorCorrectionLevel.L);
                                    }
                                    else if (value == ZXing.QrCode.Internal.ErrorCorrectionLevel.M.Name)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.QrCode.Internal.ErrorCorrectionLevel.M);
                                    }
                                    else if (value == ZXing.QrCode.Internal.ErrorCorrectionLevel.Q.Name)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.QrCode.Internal.ErrorCorrectionLevel.Q);
                                    }
                                    else if (value == ZXing.QrCode.Internal.ErrorCorrectionLevel.H.Name)
                                    {
                                        barCode.Options.Hints.Add(_feature, ZXing.QrCode.Internal.ErrorCorrectionLevel.H);
                                    }
                                }
                            }
                            break;

                        /* Specifies whether the data should be encoded to the GS1 standard type: System.Boolean, or "true" or "false" System.String value
                         */
                        case EncodeHintType.GS1_FORMAT: //bool
                            {
                                _feature = EncodeHintType.GS1_FORMAT;
                                if (bool.TryParse(value, out bool param))
                                {
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Specifies the height of the barcode image type: System.Int32
                         */
                        //HEIGHT

                        /* Specifies margin, in pixels, to use when generating the barcode. The meaning can vary by format; for example it controls margin
                         * before and after the barcode horizontally for most 1D formats. type: System.Int32, or System.String representation of the integer value
                         */
                        case EncodeHintType.MARGIN:  //int
                            {
                                _feature = EncodeHintType.MARGIN;
                                if (int.TryParse(value, out int param))
                                {
                                    if (param < 0) param = 0;
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Specifies a maximum barcode size (type ZXing.Dimension). Only applicable to Data Matrix now.
                         */
                        //MAX_SIZE

                        /* Specifies a minimum barcode size (type ZXing.Dimension). Only applicable to Data Matrix now.
                         */
                        //MIN_SIZE

                        /* Specifies the aspect ratio to use. Default is 4. type: ZXing.PDF417.Internal.PDF417AspectRatio, or 1-4.
                         */
                        case EncodeHintType.PDF417_ASPECT_RATIO: //int [1, 4]
                            {
                                _feature = EncodeHintType.PDF417_ASPECT_RATIO;
                                if (int.TryParse(value, out int param))
                                {
                                    if (param < 1) param = 1;
                                    else if (param > 4) param = 4;
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Specifies whether to use compact mode for PDF417 type: System.Boolean, or "true" or "false" System.String value
                         */
                        case EncodeHintType.PDF417_COMPACT: //bool
                            {
                                _feature = EncodeHintType.PDF417_COMPACT;
                                if (bool.TryParse(value, out bool param))
                                {
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Specifies what compaction mode to use for PDF417. type: ZXing.PDF417.Internal.Compaction or System.String value of one of its enum values
                         */
                        case EncodeHintType.PDF417_COMPACTION: //string [AUTO,BYTE,NUMERIC,TEXT]
                            {
                                _feature = EncodeHintType.PDF417_COMPACTION;
                                if (Enum.TryParse<ZXing.PDF417.Internal.Compaction>(value, out var param))
                                {
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /*Specifies the minimum and maximum number of rows and columns for PDF417. type: ZXing.PDF417.Internal.Dimensions
                         */
                        //PDF417_DIMENSIONS

                        /* Don't put the content string into the output image. type: System.Boolean
                         */
                        case EncodeHintType.PURE_BARCODE: //bool
                            {
                                _feature = EncodeHintType.QR_VERSION;
                                if (bool.TryParse(value, out bool param))
                                {
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Specifies the exact version of QR code to be encoded. (Type System.Int32, or System.String representation of the integer value).
                         */
                        case EncodeHintType.QR_VERSION: //int [1, 40] ???
                            {
                                _feature = EncodeHintType.QR_VERSION;
                                if (int.TryParse(value, out int param))
                                {
                                    barCode.Options.Hints.Add(_feature, param);
                                }
                            }
                            break;

                        /* Specifies the width of the barcode image type: System.Int32
                         */
                        //WIDTH

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

        private DataTable LoadCsvLabel(string fileName, int codePage, bool createColumnsNames = false)
        {
            DataTable table = new DataTable();
            List<string> inputStr = new List<string>();
            try
            {
                inputStr.AddRange(File.ReadAllLines(fileName, Encoding.GetEncoding(codePage)));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file:" + fileName + " : " + ex.Message);
                return table;
            }

            if (inputStr.Count <= 0) return null;
            //read headers
            int n = 0;
            if (createColumnsNames == true)
            {
                table.Columns.Clear();
                //create and count columns and read headers
                if (inputStr[0].Length != 0)
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
                        table.Columns.Add(i.ToString());
                        string tmp1 = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').Trim();
                        if (tmp1 != "") row[i] = tmp1;
                    }
                    table.Rows.Add(row);
                    n++;
                }
            }

            //read CSV content string by string
            for (; n < inputStr.Count; n++)
            {
                if (inputStr[n].ToString().Replace(Properties.Settings.Default.CSVdelimiter, ' ').Trim().TrimStart('\r').TrimStart('\n').TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').TrimEnd('\r').TrimEnd('\n') != "")
                {
                    string[] cells = inputStr[n].ToString().Split(Properties.Settings.Default.CSVdelimiter);
                    DataRow row = table.NewRow();
                    for (int i = 0; i < cells.Length - 1; i++)
                    {
                        //row[i] = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').Trim();
                        //string tmp1 = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').Trim();
                        //if (tmp1 != "") row[i] = tmp1;
                        row[i] = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r').Trim();
                    }
                    table.Rows.Add(row);
                }
            }
            return table;
        }

        private List<Template> LoadCsvTemplate(string fileName, int codePage)
        {
            List<Template> tmpLabel = new List<Template>();
            string[] inputStr = File.ReadAllLines(fileName, Encoding.GetEncoding(codePage));
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
                        for (int n = 0; n < Enum.GetNames(typeof(LabelObject)).Length; n++)
                        {
                            if (Enum.GetName(typeof(LabelObject), n) == cells[0]) templ.objectType = (LabelObject)n;
                        }
                        if (i == 0 && templ.objectType != LabelObject.label)
                        {
                            MessageBox.Show("[Line " + i.ToString() + "] Label object must be defined first");
                            return Label;
                        }
                        // label; 1 [bgColor]; 2 [objectColor]; 3 width; 4 height;
                        if (templ.objectType == LabelObject.label)
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
                            if (templ.codePage == 0) templ.codePage = codePage;
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
                        else if (templ.objectType == LabelObject.text)
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
                                else if (templ.posX >= tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = tmpLabel[0].height - 1;
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

                                FontStyle.TryParse(cells[8], out templ.fontStyle);
                                //if (templ.fontStyle) templ.fontStyle != 0 && templ.fontStyle != 1 && templ.fontStyle != 2 && templ.fontStyle != 4 && templ.fontStyle != 8)
                                if (!Enum.IsDefined(typeof(FontStyle), templ.fontStyle))
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
                        else if (templ.objectType == LabelObject.picture)
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
                                else if (templ.posX >= tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = tmpLabel[0].height - 1;
                                }

                                float.TryParse(cells[4], out templ.rotate);

                                templ.content = cells[5];
                                if (!File.Exists(path + templ.content))
                                //if (!File.Exists(@templ.content))
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

                                byte.TryParse(cells[8], out byte t);
                                templ.transparent = (t > 0);
                            }
                            else
                            {
                                MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                            }
                        }
                        // barcode; 1 [bgColor]; 2 [objectColor]; 3 posX; 4 posY; 5 [rotate]; 6 [default_data]; 7 width; 8 height; 9 bcFormat; 10 [transparent]; 11 [additional_features]
                        else if (templ.objectType == LabelObject.barcode)
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
                                else if (templ.posX >= tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[4], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = tmpLabel[0].height - 1;
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

                                BarcodeFormat.TryParse(cells[9], out templ.barCodeFormat);
                                if (!Enum.IsDefined(typeof(BarcodeFormat), templ.barCodeFormat))
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect barcode type: " + templ.barCodeFormat.ToString());
                                    templ.barCodeFormat = BarcodeFormat.QR_CODE;
                                }

                                byte.TryParse(cells[10], out byte t);
                                templ.transparent = (t > 0);

                                if (cells.Count >= 12 && cells[11].Contains("="))
                                {
                                    if (bcFeatures.Contains(cells[11].Substring(0, cells[11].IndexOf('=')).Trim()))
                                    {
                                        templ.addFeature = cells[11];
                                    }
                                    else templ.addFeature = "";
                                }
                                else templ.addFeature = "";
                            }
                            else
                            {
                                MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                            }
                        }
                        // line_coord; 1 [objectColor]; 2 posX; 3 posY; 4 ------- ; 5 [lineWidth]; 6 endX; 7 endY;
                        else if (templ.objectType == LabelObject.line_coord)
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
                                else if (templ.posX >= tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = tmpLabel[0].height - 1;
                                }

                                float.TryParse(cells[4], out templ.lineWidth);
                                if (templ.lineWidth < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect line width: " + templ.lineWidth.ToString());
                                    templ.lineWidth = 0;
                                }

                                float.TryParse(cells[5], out templ.width);
                                if (templ.width < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.width.ToString());
                                    templ.width = 0;
                                }

                                float.TryParse(cells[6], out templ.height);
                                if (templ.height < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect width: " + templ.height.ToString());
                                    templ.height = 0;
                                }
                            }
                            else
                            {
                                MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                            }
                        }
                        // line_length; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 lineLength;
                        else if (templ.objectType == LabelObject.line_length)
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
                                else if (templ.posX >= tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = tmpLabel[0].height - 1;
                                }

                                float.TryParse(cells[5], out templ.lineWidth);
                                if (templ.lineWidth < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Incorrect line width: " + templ.lineWidth.ToString());
                                    templ.lineWidth = 0;
                                }

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
                                MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                            }
                        }
                        // rectangle; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [fill];
                        else if (templ.objectType == LabelObject.rectangle)
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
                                else if (templ.posX >= tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = tmpLabel[0].height - 1;
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

                                byte.TryParse(cells[8], out byte t);
                                templ.transparent = (t > 0);
                            }
                            else
                            {
                                MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                            }
                        }
                        // ellipse; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [fill];
                        else if (templ.objectType == LabelObject.ellipse)
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
                                else if (templ.posX >= tmpLabel[0].width)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] X position out of label bounds: " + templ.posX.ToString());
                                    templ.posX = tmpLabel[0].width - 1;
                                }

                                float.TryParse(cells[3], out templ.posY);
                                if (templ.posY < 0)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = 0;
                                }
                                else if (templ.posY >= tmpLabel[0].height)
                                {
                                    MessageBox.Show("[Line " + i.ToString() + "] Y position out of label bounds: " + templ.posY.ToString());
                                    templ.posY = tmpLabel[0].height - 1;
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

                                byte.TryParse(cells[8], out byte t);
                                templ.transparent = (t > 0);
                            }
                            else
                            {
                                MessageBox.Show("[Line " + i.ToString() + "] Incomplete parameters:\r\n" + inputStr[i]);
                            }
                        }
                        tmpLabel.Add(templ);
                    }
                    else MessageBox.Show("[Line " + i.ToString() + "] Incorrect object definition");
                }
            }
            if (tmpLabel.Count > 0) return tmpLabel;
            else return Label;
        }

        private bool SaveTemplateToCSV(string fileName, List<Template> _label, int codePage)
        {
            StringBuilder output = new StringBuilder();
            char div = Properties.Settings.Default.CSVdelimiter;
            for (int i = 0; i < _label.Count; i++)
            {
                // label; 1 [bgColor]; 2 [objectColor]; 3 width; 4 height;
                if (_label[i].objectType == LabelObject.label)
                {
                    output.AppendLine(Enum.GetName(typeof(LabelObject), _label[i].objectType) + div +
                        _label[i].bgColor.Name.ToString() + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].width.ToString() + div +
                        _label[i].height.ToString() + div +
                        _label[i].dpi.ToString() + div +
                        _label[i].codePage.ToString() + div);
                }
                // text; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [default_text]; 6 fontName; 7 fontSize; 8 [fontStyle];
                else if (_label[i].objectType == LabelObject.text)
                {
                    output.AppendLine(Enum.GetName(typeof(LabelObject), _label[i].objectType) + div +
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
                else if (_label[i].objectType == LabelObject.picture)
                {
                    output.AppendLine(Enum.GetName(typeof(LabelObject), _label[i].objectType) + div +
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
                else if (_label[i].objectType == LabelObject.barcode)
                {
                    output.AppendLine(Enum.GetName(typeof(LabelObject), _label[i].objectType) + div +
                        _label[i].bgColor.Name.ToString() + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].posX.ToString() + div +
                        _label[i].posY.ToString() + div +
                        _label[i].rotate.ToString() + div +
                        _label[i].content.ToString() + div +
                        _label[i].width.ToString() + div +
                        _label[i].height.ToString() + div +
                        _label[i].barCodeFormat.ToString() + div +
                        Convert.ToInt32(_label[i].transparent).ToString() + div +
                        _label[i].addFeature.ToString() + div);
                }
                // line_coord; 1 [objectColor]; 2 posX; 3 posY; 4 [lineWidth]; 5 endX; 6 endY;
                else if (_label[i].objectType == LabelObject.line_coord)
                {
                    output.AppendLine(Enum.GetName(typeof(LabelObject), _label[i].objectType) + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].posX.ToString() + div +
                        _label[i].posY.ToString() + div +
                        _label[i].lineWidth.ToString() + div +
                        _label[i].width.ToString() + div +
                        _label[i].height.ToString() + div);
                }
                // line_length; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 lineLength;
                else if (_label[i].objectType == LabelObject.line_length)
                {
                    output.AppendLine(Enum.GetName(typeof(LabelObject), _label[i].objectType) + div +
                        _label[i].fgColor.Name.ToString() + div +
                        _label[i].posX.ToString() + div +
                        _label[i].posY.ToString() + div +
                        _label[i].rotate.ToString() + div +
                        _label[i].lineWidth.ToString() + div +
                        _label[i].lineLength.ToString() + div);
                }
                // rectangle; 1 [objectColor]; 2 posX; 3 posY; 4 [rotate]; 5 [lineWidth]; 6 width; 7 height; 8 [transparent];
                else if (_label[i].objectType == LabelObject.rectangle)
                {
                    output.AppendLine(Enum.GetName(typeof(LabelObject), _label[i].objectType) + div +
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
                else if (_label[i].objectType == LabelObject.ellipse)
                {
                    output.AppendLine(Enum.GetName(typeof(LabelObject), _label[i].objectType) + div +
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
                File.WriteAllBytes(fileName, Encoding.GetEncoding(codePage).GetBytes(output.ToString()));
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error writing to file " + SaveFileDialog1.FileName + ": " + ex.Message);
                err = false;
            }
            return err;
        }

        private bool SaveTableToCSV(string fileName, DataTable dataTable, bool saveColumnNames, char csvDivider = ';', int codePage = -1)
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
                    output.Append(dataTable.Rows[j].ItemArray[i].ToString() + csvDivider);
                }
                output.AppendLine();
            }

            bool err = true;
            try
            {
                File.WriteAllBytes(fileName, Encoding.GetEncoding(codePage).GetBytes(output.ToString()));
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
            if (lineNumber >= dataTable.Rows.Count || _label.Count <= 0) return null;
            if (_label[0].objectType != LabelObject.label)
            {
                MessageBox.Show("1st object must be \"label\"");
                return null;
            }
            currentObjectsPositions.Clear();
            for (int i = 0; i < _label.Count; i++)
            {
                currentObjectsPositions.Add(new Rectangle());
                if (_label[i].objectType == LabelObject.label)
                {
                    Color bColor = _label[i].bgColor;
                    float width = _label[i].width;
                    float height = _label[i].height;

                    img = new Bitmap((int)width, (int)height, PixelFormat.Format32bppPArgb);
                    FillBackground(img, bColor, width, height);
                    //pictureBox_label.Image = img;
                }
                else if (_label[i].objectType == LabelObject.text)
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
                    FontStyle fontStyle = _label[i].fontStyle;

                    currentObjectsPositions[i] = DrawText(img, fColor, posX, posY, content, fontname, fontSize, rotate, fontStyle);
                }
                else if (_label[i].objectType == LabelObject.picture)
                {
                    Color fColor = _label[i].fgColor;
                    float posX = _label[i].posX;
                    float posY = _label[i].posY;

                    string content = path + _label[i].content;
                    if (lineNumber > -1)
                        if (dataTable.Rows[lineNumber].ItemArray[i - 1].ToString() != "")
                            content = path + dataTable.Rows[lineNumber].ItemArray[i - 1].ToString();

                    float rotate = _label[i].rotate;
                    float width = _label[i].width;
                    float height = _label[i].height;
                    bool transparent = _label[i].transparent;

                    currentObjectsPositions[i] = DrawPicture(img, fColor, posX, posY, content, rotate, width, height, transparent);
                }
                else if (_label[i].objectType == LabelObject.barcode)
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

                    BarcodeFormat BCformat = _label[i].barCodeFormat;
                    float rotate = _label[i].rotate;
                    string feature = _label[i].addFeature;
                    bool transparent = _label[i].transparent;

                    currentObjectsPositions[i] = DrawBarcode(img, bColor, fColor, posX, posY, width, height, content, BCformat, rotate, feature, transparent);
                }
                else if (_label[i].objectType == LabelObject.line_coord)
                {
                    Color fColor = _label[i].fgColor;
                    float posX = _label[i].posX;
                    float posY = _label[i].posY;
                    float rotate = _label[i].rotate;
                    float lineWidth = _label[i].lineWidth;
                    float endX = _label[i].width;
                    float endY = _label[i].height;
                    currentObjectsPositions[i] = DrawLineCoord(img, fColor, posX, posY, endX, endY, lineWidth);
                }
                else if (_label[i].objectType == LabelObject.line_length)
                {
                    Color fColor = _label[i].fgColor;
                    float posX = _label[i].posX;
                    float posY = _label[i].posY;
                    float rotate = _label[i].rotate;
                    float lineWidth = _label[i].lineWidth;
                    float length = _label[i].lineLength;
                    currentObjectsPositions[i] = DrawLineLength(img, fColor, posX, posY, length, rotate, lineWidth);
                }
                else if (_label[i].objectType == LabelObject.rectangle)
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
                else if (_label[i].objectType == LabelObject.ellipse)
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
                    MessageBox.Show("Incorrect object: " + _label[i].objectType);
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
                templ.objectType = (LabelObject)comboBox_object.SelectedIndex;
                if (Label[0].objectType == LabelObject.label)
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
                templ.dpi = 300;

                int cp = Properties.Settings.Default.CodePage;
                if (int.TryParse(comboBox_encoding.SelectedItem.ToString().Substring(0, comboBox_encoding.SelectedItem.ToString().IndexOf('=')), out cp)) Properties.Settings.Default.CodePage = cp;
                templ.codePage = cp;

                templ.rotate = 0;
                templ.content = "";
                templ.width = 1;
                templ.height = 1;
                templ.transparent = true;
                templ.barCodeFormat = BarcodeFormat.AZTEC;
                templ.fontSize = 10;
                templ.fontStyle = FontStyle.Regular;
                templ.fontName = "Arial";
                templ.addFeature = "";
                templ.lineLength = 1;
                templ.lineWidth = 1;
            }
            else
            {
                templ = Label[n];
                // label; [bgColor]; [objectColor]; width; height;
                if (templ.objectType == LabelObject.label)
                {
                    if (comboBox_backgroundColor.SelectedItem.ToString() == "Background color") templ.bgColor = Label[0].bgColor;
                    else templ.bgColor = Color.FromName(comboBox_backgroundColor.SelectedItem.ToString());

                    if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                    else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                    float.TryParse(textBox_width.Text, out float f);
                    templ.width = f * mult;

                    float.TryParse(textBox_height.Text, out f);
                    templ.height = f * mult;
                }

                // text; [objectColor]; posX; posY; [rotate]; [default_text]; fontName; fontSize; [fontStyle];
                else if (templ.objectType == LabelObject.text)
                {
                    if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.bgColor = Label[0].fgColor;
                    else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                    float.TryParse(textBox_posX.Text, out float f);
                    templ.posX = f * mult;

                    float.TryParse(textBox_posY.Text, out f);
                    templ.posY = f * mult;

                    float.TryParse(textBox_rotate.Text, out f);
                    templ.rotate = f;

                    templ.content = textBox_content.Text;

                    textBox_fontSize.Text = Evaluate(textBox_fontSize.Text).ToString();
                    float.TryParse(textBox_fontSize.Text, out float b);
                    templ.fontSize = b;

                    FontStyle.TryParse(comboBox_fontStyle.SelectedItem.ToString().Substring(0, 1), out FontStyle s);
                    templ.fontStyle = s;

                    templ.fontName = comboBox_fontName.SelectedItem.ToString();
                }

                // picture; [objectColor]; posX; posY; [rotate]; [default_file]; [width]; [height]; [transparent];
                else if (templ.objectType == LabelObject.picture)
                {
                    if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.bgColor = Label[0].fgColor;
                    else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                    float.TryParse(textBox_posX.Text, out float f);
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
                else if (templ.objectType == LabelObject.barcode)
                {
                    if (comboBox_backgroundColor.SelectedItem.ToString() == "Background color") templ.bgColor = Label[0].bgColor;
                    else templ.bgColor = Color.FromName(comboBox_backgroundColor.SelectedItem.ToString());

                    if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                    else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                    float.TryParse(textBox_posX.Text, out float f);
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

                    BarcodeFormat i = BarcodeFormat.QR_CODE;
                    BarcodeFormat.TryParse(comboBox_fontName.SelectedItem.ToString(), out i);
                    templ.barCodeFormat = i;

                    if (comboBox_fontStyle.SelectedItem.ToString() != "") templ.addFeature = comboBox_fontStyle.SelectedItem.ToString() + "=" + textBox_fontSize.Text;
                    else templ.addFeature = "";
                }

                // line_coord; [objectColor]; posX; posY; --- ; [lineWidth]; endX; endY; (lineLength = -1)
                else if (templ.objectType == LabelObject.line_coord)
                {
                    if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                    else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                    float.TryParse(textBox_posX.Text, out float f);
                    templ.posX = f * mult;

                    float.TryParse(textBox_posY.Text, out f);
                    templ.posY = f * mult;

                    textBox_fontSize.Text = Evaluate(textBox_fontSize.Text).ToString();
                    float.TryParse(textBox_fontSize.Text, out f);
                    templ.lineWidth = f * mult;

                    float.TryParse(textBox_width.Text, out f);
                    templ.width = f * mult;

                    float.TryParse(textBox_height.Text, out f);
                    templ.height = f * mult;
                }

                // line_length; [objectColor]; posX; posY; [rotate]; [lineWidth]; lineLength;
                else if (templ.objectType == LabelObject.line_length)
                {
                    if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                    else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                    float.TryParse(textBox_posX.Text, out float f);
                    templ.posX = f * mult;

                    float.TryParse(textBox_posY.Text, out f);
                    templ.posY = f * mult;

                    textBox_fontSize.Text = Evaluate(textBox_fontSize.Text).ToString();
                    float.TryParse(textBox_fontSize.Text, out f);
                    templ.lineWidth = f * mult;

                    float.TryParse(textBox_rotate.Text, out f);
                    templ.rotate = f;

                    textBox_content.Text = Evaluate(textBox_content.Text).ToString();
                    float.TryParse(textBox_content.Text, out f);
                    templ.lineLength = f * mult;
                }

                // rectangle; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
                else if (templ.objectType == LabelObject.rectangle)
                {
                    if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                    else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                    float.TryParse(textBox_posX.Text, out float f);
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
                else if (templ.objectType == LabelObject.ellipse)
                {
                    if (comboBox_objectColor.SelectedItem.ToString() == "Default object color") templ.fgColor = Label[0].fgColor;
                    else templ.fgColor = Color.FromName(comboBox_objectColor.SelectedItem.ToString());

                    float.TryParse(textBox_posX.Text, out float f);
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
                comboBox_object.SelectedItem = Label[n].objectType;
                // label; [bgColor]; [objectColor]; width; height;
                if (Label[n].objectType == LabelObject.label)
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
                else if (Label[n].objectType == LabelObject.text)
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
                    //comboBox_fontStyle.Items.AddRange(Enum.GetNames(typeof(FontStyle)));
                    comboBox_fontStyle.Items.AddRange(GetFontStyleList());
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
                    textBox_fontSize.Text = (Label[n].fontSize).ToString("F4");
                }

                // picture; [objectColor]; posX; posY; [rotate]; [default_file]; [width]; [height]; [transparent];
                else if (Label[n].objectType == LabelObject.picture)
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
                else if (Label[n].objectType == LabelObject.barcode)
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
                    str = Label[n].barCodeFormat.ToString();
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

                    str = Label[n].addFeature;
                    if (str.Contains("="))
                    {
                        for (int i = 0; i < comboBox_fontStyle.Items.Count; i++)
                        {
                            if (str.Substring(0, str.IndexOf('=')) == comboBox_fontStyle.Items[i].ToString())
                            {
                                comboBox_fontStyle.SelectedIndex = i;
                                break;
                            }
                        }
                        textBox_fontSize.Enabled = true;
                        label_fontSize.Text = "Feature value";
                        textBox_fontSize.Text = Label[n].addFeature.Substring(Label[n].addFeature.IndexOf('=') + 1);
                    }
                    checkBox_fill.Enabled = true;
                    checkBox_fill.Text = "Transparent";
                    checkBox_fill.Checked = Label[n].transparent;
                }

                // line_coord; [objectColor]; posX; posY; --- ; [lineWidth]; endX; endY; (lineLength = -1)
                else if (Label[n].objectType == LabelObject.line_coord)
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

                    textBox_rotate.Enabled = false;
                    label_rotate.Text = "Rotate";
                    textBox_rotate.Text = "";

                    textBox_content.Enabled = false;
                    label_content.Text = "Line length (empty to use coordinates)";
                    textBox_content.Text = ""; //-V3024
                }

                // line_length; [objectColor]; posX; posY; [rotate]; [lineWidth]; lineLength;
                else if (Label[n].objectType == LabelObject.line_length)
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

                    textBox_width.Enabled = false;
                    label_width.Text = "endX";
                    textBox_width.Text = (Label[n].width / mult).ToString("F4");

                    textBox_height.Enabled = false;
                    label_height.Text = "endY";
                    textBox_height.Text = (Label[n].height / mult).ToString("F4");

                    textBox_rotate.Enabled = true;
                    label_rotate.Text = "Rotate";
                    textBox_rotate.Text = Label[n].rotate.ToString();

                    textBox_content.Enabled = true;
                    label_content.Text = "Line length (empty to use coordinates)";
                    textBox_content.Text = (Label[n].lineLength / mult).ToString("F4"); //-V3024
                }

                // rectangle; [objectColor]; posX; posY; [rotate]; [lineWidth]; width; height; [fill];
                else if (Label[n].objectType == LabelObject.rectangle)
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
                else if (Label[n].objectType == LabelObject.ellipse)
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
                objectList.Add(i.ToString() + " " + t.objectType);
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

        private string[] GetFontStyleList()
        {
            return Enum.GetNames(typeof(FontStyle));
        }

        private string[] GetBarcodeList()
        {
            return Enum.GetNames(typeof(BarcodeFormat));
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
                dataGridView_labels.DataSource = null;
                LabelsDatabase.Clear();
                LabelsDatabase.Columns.Clear();
                LabelsDatabase.Rows.Clear();
                textBox_labelsName.Clear();
                //Label.Clear();
                path = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.LastIndexOf('\\') + 1);
                Label = LoadCsvTemplate(openFileDialog1.FileName, Label[0].codePage);

                TextBox_dpi_Leave(this, EventArgs.Empty);
                button_importLabels.Enabled = true;
                LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
                pictureBox_label.Image = LabelBmp;

                textBox_templateName.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\') + 1);
                //save path to template
                //create colums and fill 1 row with default values
                for (int i = 1; i < Label.Count; i++)
                {
                    LabelsDatabase.Columns.Add((i).ToString() + " " + Label[i].objectType);
                }
                DataRow row = LabelsDatabase.NewRow();
                for (int i = 1; i < Label.Count; i++)
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

                path = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.LastIndexOf('\\') + 1);
                LabelsDatabase = LoadCsvLabel(openFileDialog1.FileName, Label[0].codePage, checkBox_columnNames.Checked);

                if (LabelsDatabase.Rows.Count > 0)
                {
                    dataGridView_labels.DataSource = LabelsDatabase;
                    foreach (DataGridViewColumn column in dataGridView_labels.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
                    //check for picture file existence
                    foreach (DataRow row in LabelsDatabase.Rows)
                    {
                        for (int i = 0; i < LabelsDatabase.Columns.Count; i++)
                        {
                            if (Label[i + 1].objectType == LabelObject.picture && !File.Exists(path + row.ItemArray[i].ToString()))
                            {
                                MessageBox.Show("[Line " + (i + 1).ToString() + "] File not exist: " + path + row.ItemArray[i].ToString());
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
                //save path to template
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
                if (!SaveTemplateToCSV(SaveFileDialog1.FileName, Label, Label[0].codePage)) MessageBox.Show("Error writing to file " + SaveFileDialog1.FileName);
            }
            else if (SaveFileDialog1.Title == "Save label data as .CSV...")
            {
                if (!SaveTableToCSV(SaveFileDialog1.FileName, LabelsDatabase, checkBox_columnNames.Checked, Properties.Settings.Default.CSVdelimiter, Label[0].codePage)) MessageBox.Show("Error writing to file " + SaveFileDialog1.FileName);
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
                pictureBox_label.Dock = DockStyle.None;
                pictureBox_label.SizeMode = PictureBoxSizeMode.Normal;
                pictureBox_label.Width = (int)Label[0].width;
                pictureBox_label.Height = (int)Label[0].height;
            }
            else
            {
                pictureBox_label.Dock = DockStyle.Fill;
                pictureBox_label.SizeMode = PictureBoxSizeMode.Zoom;
            }
        }

        private void DataGridView_labels_SelectionChanged(object sender, EventArgs e)
        {
            if (!cmdLinePrint && dataGridView_labels.CurrentCell.RowIndex < LabelsDatabase.Rows.Count)
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
                    //create column headers
                    LabelsDatabase.Columns.Clear();

                    if (Label.Count > 1)
                    {
                        //create and count columns and read headers
                        for (int i = 1; i < Label.Count; i++)
                        {
                            LabelsDatabase.Columns.Add(i.ToString() + " " + Label[i].objectType);
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
                        foreach (DataRow row in LabelsDatabase.Rows)
                        {
                            for (int i = 0; i < LabelsDatabase.Columns.Count; i++)
                            {
                                if (Label[i + 1].objectType == LabelObject.picture && !File.Exists(path + row.ItemArray[i].ToString()))
                                {
                                    MessageBox.Show("[Line " + (i + 1).ToString() + "] File not exist: " + path + row.ItemArray[i].ToString());
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
                    }
                    else
                    {
                        GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
                        pictureBox_label.Image = LabelBmp;
                    }
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

                listBox_objectsMulti.SelectedIndex = listBox_objects.SelectedIndex;

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
            pictureBox_label.Width = LabelBmp.Width;
            pictureBox_label.Height = LabelBmp.Height;
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
            Template templ = Label[0];
            int cp = Properties.Settings.Default.CodePage;
            if (int.TryParse(comboBox_encoding.SelectedItem.ToString().Substring(0, comboBox_encoding.SelectedItem.ToString().IndexOf('=')), out cp)) Properties.Settings.Default.CodePage = cp;
            templ.codePage = cp;
            Label[0] = templ;
        }

        private void TextBox_dpi_Leave(object sender, EventArgs e)
        {
            Template templ = Label[0];
            float diff = 0;

            float.TryParse(textBox_dpi.Text, out templ.dpi);
            diff = templ.dpi / Label[0].dpi;
            Label[0] = templ;
            textBox_dpi.Text = Label[0].dpi.ToString();
            units[1] = (float)(Label[0].dpi / 25.4);
            units[2] = (float)(Label[0].dpi / 2.54);
            units[3] = Label[0].dpi;
            mult = units[comboBox_units.SelectedIndex];
            //recalculate all objects
            for (int i = 0; i < Label.Count; i++)
            {
                Template tmp = Label[i];
                tmp.width *= diff;
                tmp.height *= diff;
                tmp.fontSize *= diff;
                tmp.lineLength *= diff;
                tmp.lineWidth *= diff;
                tmp.posX *= diff;
                tmp.posY *= diff;
                Label[i] = tmp;
            }
            LabelBmp = GenerateLabel(Label, LabelsDatabase, -1, LabelBmp);
            pictureBox_label.Image = LabelBmp;
            pictureBox_label.Width = LabelBmp.Width;
            pictureBox_label.Height = LabelBmp.Height;
            _templateChanged = true;
            ShowObjectInGUI(listBox_objects.SelectedIndex);
        }

        private void ComboBox_units_SelectedIndexChanged(object sender, EventArgs e)
        {
            mult = units[comboBox_units.SelectedIndex];
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

                    if (Label[n].objectType == LabelObject.line_coord || Label[n].objectType == LabelObject.line_length)
                        DrawSelection(LabelBmp, _borderColor, currentObjectsPositions[n].X, currentObjectsPositions[n].Y, currentObjectsPositions[n].Width, currentObjectsPositions[n].Height, 0, 1);
                    else
                        DrawSelection(LabelBmp, _borderColor, currentObjectsPositions[n].X, currentObjectsPositions[n].Y, currentObjectsPositions[n].Width, currentObjectsPositions[n].Height, Label[n].rotate, 1);
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
            var saveSelections = listBox_objectsMulti.SelectedIndices;
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1)
                {
                    Template templ = Label[n];
                    float y = (float)numericUpDown_scale.Value;
                    templ.posY -= y;
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
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
                    float x = (float)numericUpDown_scale.Value;
                    templ.posX -= x;
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
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
                    float y = (float)numericUpDown_scale.Value;
                    templ.posY += y;
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
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
                    float x = (float)numericUpDown_scale.Value;
                    templ.posX += x;
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_enlarge_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1)
                {
                    Template templ = Label[n];
                    float x = (float)numericUpDown_scale.Value;
                    if (templ.objectType == LabelObject.text)
                    {
                        templ.fontSize += x;
                    }
                    else
                    {
                        templ.width += x;
                        templ.height += x;
                    }
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_reduce_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1)
                {
                    Template templ = Label[n];
                    float x = (float)numericUpDown_scale.Value;
                    if (templ.objectType == LabelObject.text)
                    {
                        templ.fontSize -= x;
                    }
                    else
                    {
                        templ.width -= x;
                        templ.height -= x;
                    }
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_scaleRight_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1 && Label[n].objectType != LabelObject.text && Label[n].objectType != LabelObject.line_length)
                {
                    Template templ = Label[n];
                    float x = (float)numericUpDown_scale.Value;
                    templ.width += x;
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_scaleLeft_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1 && Label[n].objectType != LabelObject.text && Label[n].objectType != LabelObject.line_length)
                {
                    Template templ = Label[n];
                    float x = (float)numericUpDown_scale.Value;
                    templ.width -= x;
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_scaleDown_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1 && Label[n].objectType != LabelObject.text && Label[n].objectType != LabelObject.line_length)
                {
                    Template templ = Label[n];
                    float x = (float)numericUpDown_scale.Value;
                    templ.height += x;
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
            foreach (int n in k) listBox_objectsMulti.SetSelected(n, true);
            _templateChanged = true;
            ShowObjectInGUI(listBox_objectsMulti.SelectedIndex);
            listBox_objectsMulti.SelectedIndexChanged += new EventHandler(ListBox_objects_SelectedIndexChanged);
        }

        private void Button_scaleUp_Click(object sender, EventArgs e)
        {
            List<int> k = new List<int>();
            foreach (int n in listBox_objectsMulti.SelectedIndices)
            {
                if (n > 0 && n < listBox_objectsMulti.Items.Count - 1 && Label[n].objectType != LabelObject.text && Label[n].objectType != LabelObject.line_length)
                {
                    Template templ = Label[n];
                    float x = (float)numericUpDown_scale.Value;
                    templ.height -= x;
                    Label[n] = templ;
                    k.Add(n);
                }
            }
            listBox_objectsMulti.SelectedIndexChanged -= new EventHandler(ListBox_objects_SelectedIndexChanged);
            listBox_objectsMulti.SelectedIndex = -1;
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

        private void TextBox_rangeFrom_Leave(object sender, EventArgs e)
        {
            int.TryParse(textBox_rangeFrom.Text, out int n);
            if (n < 1) textBox_rangeFrom.Text = "1";
            if (n >= 1) textBox_rangeFrom.Text = n.ToString();
        }

        private void TextBox_rangeTo_Leave(object sender, EventArgs e)
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
            if (listBox_objectsMulti.SelectedIndices.Count <= 0 || listBox_objectsMulti.SelectedIndices.Contains(0))
            {
                button_moveUp.Enabled = false;
                button_moveDown.Enabled = false;
                button_moveLeft.Enabled = false;
                button_moveRight.Enabled = false;
            }
            else
            {
                button_moveUp.Enabled = true;
                button_moveDown.Enabled = true;
                button_moveLeft.Enabled = true;
                button_moveRight.Enabled = true;
            }

            if (listBox_objectsMulti.SelectedIndices.Count <= 0)
            {
                button_moveUp.Enabled = false;
                button_moveDown.Enabled = false;
                button_moveLeft.Enabled = false;
                button_moveRight.Enabled = false;
            }
            else
            {
                button_moveUp.Enabled = true;
                button_moveDown.Enabled = true;
                button_moveLeft.Enabled = true;
                button_moveRight.Enabled = true;
            }
        }

        private void Button_printCurrent_Click(object sender, EventArgs e)
        {
            if (dataGridView_labels.CurrentRow.Index >= LabelsDatabase.Rows.Count) return;
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
            int _pagesTo = LabelsDatabase.Rows.Count;

            int.TryParse(textBox_rangeFrom.Text, out int _pagesFrom);
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

        //show current mouse cursor position on the picture
        private void PictureBox_label_MouseMove(object sender, MouseEventArgs e)
        {
            if (pictureBox_label.Image == null) return;
            if (checkBox_scale.Checked)
            {
                mouseX = e.X;
                mouseY = e.Y;
            }
            else
            {
                float scaleW = (float)pictureBox_label.Width / LabelBmp.Width;
                float scaleH = (float)pictureBox_label.Height / LabelBmp.Height;
                if (scaleW > scaleH)
                {
                    mouseY = e.Y / scaleH;
                    float corr = (pictureBox_label.Width - (LabelBmp.Width * scaleH)) / 2;
                    mouseX = (e.X - corr) / scaleH;
                }
                else
                {
                    mouseX = e.X / scaleW;
                    float corr = (pictureBox_label.Height - LabelBmp.Height * scaleW) / 2;
                    mouseY = (e.Y - corr) / scaleW;
                }
            }
            textBox_mX.Text = "X=" + (mouseX / mult).ToString("F4") + comboBox_units.SelectedItem.ToString();
            textBox_mX.Text += "; Y=" + (mouseY / mult).ToString("F4") + comboBox_units.SelectedItem.ToString();
            int n = LabelObjectPointed((int)mouseX, (int)mouseY);
            if (n > 0) textBox_mX.Text += "; [" + n + "] " + Label[n].objectType;
        }

        private bool IsInPolygon(PointF[] poly, PointF pnt)
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
        }

        private PointF[] RotatePolygon(PointF zeroPoint, PointF[] poly, float angleDegree)
        {
            PointF[] p = new PointF[poly.Length];
            for (int i = 0; i < poly.Length; i++)
            {
                p[i] = RotateLine(zeroPoint, poly[i], angleDegree);
            }
            return p;
        }

        private PointF RotateLine(PointF center, PointF end, double angleDeg)
        {
            PointF result = new PointF();
            if (angleDeg % 360 == 0) return end;
            double angleRad = (angleDeg * Math.PI) / 180;
            result.X = (float)((end.X - center.X) * Math.Cos(angleRad) - (end.Y - center.Y) * Math.Sin(angleRad) + center.X);
            result.Y = (float)((end.X - center.X) * Math.Sin(angleRad) + (end.Y - center.Y) * Math.Cos(angleRad) + center.Y);
            return result;
        }

        //select object with click
        private void PictureBox_label_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex != 1 && tabControl1.SelectedIndex != 2) return;

            int n = LabelObjectPointed((int)mouseX, (int)mouseY);
            if (n > 0)
            {
                if (tabControl1.SelectedIndex == 1) listBox_objects.SelectedIndex = n;
                else if (tabControl1.SelectedIndex == 2)
                {
                    listBox_objectsMulti.SetSelected(n, !listBox_objectsMulti.GetSelected(n));
                }
            }
        }

        public int LabelObjectPointed(int posX, int posY)
        {
            for (int i = 0; i < Label.Count; i++)
            {
                PointF[] rect = new PointF[4];
                rect[0].X = currentObjectsPositions[i].X;
                rect[0].Y = currentObjectsPositions[i].Y;

                rect[1].X = currentObjectsPositions[i].X + currentObjectsPositions[i].Width;
                rect[1].Y = currentObjectsPositions[i].Y;

                rect[2].X = currentObjectsPositions[i].X + currentObjectsPositions[i].Width;
                rect[2].Y = currentObjectsPositions[i].Y + currentObjectsPositions[i].Height;

                rect[3].X = currentObjectsPositions[i].X;
                rect[3].Y = currentObjectsPositions[i].Y + currentObjectsPositions[i].Height;

                if (Label[i].objectType != LabelObject.line_coord)
                {
                    rect = RotatePolygon(rect[0], rect, Label[i].rotate);
                }

                if (IsInPolygon(rect, new PointF(posX, posY)))
                {
                    return i;
                }
            }
            return -1;
        }

        #endregion

    }
}

