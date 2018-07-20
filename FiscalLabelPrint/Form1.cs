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

namespace FiscalLabelPrint
{
    public partial class Form1 : Form
    {
        enum LabelTtype
        {
            label,
            text,
            picture,
            barcode,
            qrcode,
        }

        string[] _labelTypes = { "label", "text", "picture", "barcode", "2dcode" };

        struct template
        {
            public string type;
            public int posX;
            public int posY;
            public int rotate;
            public string content;
            public int width;
            public int height;
            public bool transparent;
            public int BCformat;
            public int fontSize;
            public byte fontStyle;
            public string fontName;
        }
        template[] Label = new template[0];

        DataTable LabelsDatabase = new DataTable();

        int labelWidth = 0;
        int labelHeight = 0;

        int pagesFrom = 0;
        int pagesTo = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void drawText(Image img, int posX, int posY, string text, string fontName, int fontSize, int rotateDeg = 0, FontStyle fontStyle = FontStyle.Regular)
        {
            /*InstalledFontCollection installedFontCollection = new InstalledFontCollection();
            foreach (FontFamily fontFamily in installedFontCollection.Families)
            {
                if (fontFamily.Name == fontName) ;
            }
            installedFontCollection.Dispose();*/

            Font textFont = new Font(fontName, fontSize, fontStyle); //creates new font
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            // Save the graphics state.
            GraphicsState state = g.Save();
            g.ResetTransform();
            // Rotate.
            g.RotateTransform(rotateDeg);
            // Translate to desired position. Be sure to append the rotation so it occurs after the rotation.
            g.TranslateTransform(posX, posY, MatrixOrder.Append);
            // Draw the text at the origin.
            g.DrawString(text, textFont, Brushes.Black, 0, 0);
            // Restore the graphics state.
            g.Restore(state);
        }

        private void drawPicture(Image img, int posX, int posY, string fileName, int rotateDeg = 0, int width = 0, int height = 0, bool makeTransparent = true)
        {
            Bitmap newPicture = new Bitmap(1, 1);
            try
            {
                newPicture = new Bitmap(@fileName, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file:" + fileName + " : " + ex.Message);
                return;
            }
            if (width == 0) width = newPicture.Width;
            if (height == 0) height = newPicture.Height;
            if (makeTransparent) newPicture.MakeTransparent(Color.White);
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.HighQuality;
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

        private void drawBarcode(Image img, int posX, int posY, int width, int height, string BCdata, BarcodeFormat bcFormat, int rotateDeg = 0)
        {
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            GraphicsState state = g.Save();
            g.ResetTransform();

            // Rotate.
            g.RotateTransform(rotateDeg);

            // Translate to desired position. Be sure to append the rotation so it occurs after the rotation.
            g.TranslateTransform(posX, posY, MatrixOrder.Append);

            var barCode = new BarcodeWriter();
            barCode.Format = bcFormat;
            barCode.Options = new EncodingOptions
            {
                PureBarcode = true,
                Height = height,
                Width = width
            };
            Bitmap newPicture = barCode.Write(BCdata);
            newPicture.MakeTransparent(Color.White);

            g.DrawImage(newPicture, 0, 0);

            // Restore the graphics state.
            g.Restore(state);
        }

        private void draw2Dcode(Image img, int posX, int posY, int width, int height, string BCdata, BarcodeFormat bcFormat, int rotateDeg = 0)
        {
            Graphics g = Graphics.FromImage(img);
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            GraphicsState state = g.Save();
            g.ResetTransform();

            // Rotate.
            g.RotateTransform(rotateDeg);

            // Translate to desired position. Be sure to append the rotation so it occurs after the rotation.
            g.TranslateTransform(posX, posY, MatrixOrder.Append);

            var barCode = new BarcodeWriter();
            barCode.Format = bcFormat;
            barCode.Options = new EncodingOptions
            {
                PureBarcode = true,
                Height = height,
                Width = width
            };
            Bitmap newPicture = barCode.Write(BCdata);
            newPicture.MakeTransparent(Color.White);

            g.DrawImage(newPicture, 0, 0);

            // Restore the graphics state.
            g.Restore(state);
        }

        public void ReadCsv(string fileName, DataTable table)
        {
            table.Clear();
            table.Columns.Clear();
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
            int c = inputFile.ReadByte();
            while (c != '\r' && c != '\n' && c != -1)
            {
                byte[] b = new byte[1];
                b[0] = (byte)c;
                inputStr.Append(Encoding.GetEncoding(LabelPrint.Properties.Settings.Default.CodePage).GetString(b));
                c = inputFile.ReadByte();
            }

            //create and count columns and read headers
            int colNum = 0;
            if (inputStr.Length != 0)
            {
                string[] cells = inputStr.ToString().Split(LabelPrint.Properties.Settings.Default.CSVdelimiter);
                colNum = cells.Length - 1;
                for (int i = 0; i < colNum; i++)
                {
                    table.Columns.Add(cells[i]);
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
                LabelsDatabase.Clear();
                ReadCsv(openFileDialog1.FileName, LabelsDatabase);
                //dataGridView_labels.Rows.Clear();
                //dataGridView_labels.Columns.Clear();
                dataGridView_labels.DataSource = LabelsDatabase;
                if (LabelsDatabase != null)
                {
                    foreach (DataGridViewColumn column in dataGridView_labels.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
                    button_printCurrent.Enabled = true;
                    button_printAll.Enabled = true;
                    button_printRange.Enabled = true;
                    textBox_rangeFrom.Text = "0";
                    textBox_rangeTo.Text = (LabelsDatabase.Rows.Count - 1).ToString();
                    setRowNumber(dataGridView_labels);
                }
                if (dataGridView_labels.Columns.Count != Label.Length - 1) MessageBox.Show("Template doesn't match label file");
                dataGridView_labels.CurrentCell = dataGridView_labels.Rows[0].Cells[0];
                dataGridView_labels.Rows[0].Selected = true;
                generateLabel(dataGridView_labels.CurrentCell.RowIndex);
                textBox_labelsName.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\') + 1);
            }
            else if (openFileDialog1.Title == "Open template .CSV file")
            {
                string[] inputStr = File.ReadAllLines(openFileDialog1.FileName);
                Label = new template[inputStr.Length];
                for (int i = 0; i < inputStr.Length; i++)
                {
                    if (inputStr[i].ToString().Trim() != "")
                    {
                        string[] cells = inputStr[i].ToString().Split(LabelPrint.Properties.Settings.Default.CSVdelimiter);
                        for (int i1 = 0; i1 < cells.Length; i1++) cells[i1] = cells[i1].Trim();

                        Label[i].type = cells[0];
                        if (cells.Length >= 3 && Label[i].type == _labelTypes[(int)LabelTtype.label])
                        {
                            int.TryParse(cells[1], out Label[i].width);
                            int.TryParse(cells[2], out Label[i].height);
                            labelWidth = Label[i].width;
                            labelHeight = Label[i].height;
                        }
                        else if (cells.Length >= 8 && Label[i].type == _labelTypes[(int)LabelTtype.text])
                        {
                            int.TryParse(cells[1], out Label[i].posX);
                            int.TryParse(cells[2], out Label[i].posY);
                            int.TryParse(cells[3], out Label[i].rotate);
                            Label[i].content = cells[4];
                            Label[i].fontName = cells[5];
                            int.TryParse(cells[6], out Label[i].fontSize);
                            byte.TryParse(cells[7], out Label[i].fontStyle);
                        }
                        else if (cells.Length >= 8 && Label[i].type == _labelTypes[(int)LabelTtype.picture])
                        {
                            int.TryParse(cells[1], out Label[i].posX);
                            int.TryParse(cells[2], out Label[i].posY);
                            int.TryParse(cells[3], out Label[i].rotate);
                            Label[i].content = cells[4];
                            int.TryParse(cells[5], out Label[i].width);
                            int.TryParse(cells[6], out Label[i].height);
                            int t = 0;
                            int.TryParse(cells[7], out t);
                            Label[i].transparent = (t > 0);
                        }
                        else if (cells.Length >= 8 && Label[i].type == _labelTypes[(int)LabelTtype.barcode])
                        {
                            int.TryParse(cells[1], out Label[i].posX);
                            int.TryParse(cells[2], out Label[i].posY);
                            int.TryParse(cells[3], out Label[i].rotate);
                            Label[i].content = cells[4];
                            int.TryParse(cells[5], out Label[i].width);
                            int.TryParse(cells[6], out Label[i].height);
                            int.TryParse(cells[7], out Label[i].BCformat);
                        }
                        else if (cells.Length >= 8 && Label[i].type == _labelTypes[(int)LabelTtype.qrcode])
                        {
                            int.TryParse(cells[1], out Label[i].posX);
                            int.TryParse(cells[2], out Label[i].posY);
                            int.TryParse(cells[3], out Label[i].rotate);
                            Label[i].content = cells[4];
                            int.TryParse(cells[5], out Label[i].width);
                            int.TryParse(cells[6], out Label[i].height);
                            int.TryParse(cells[7], out Label[i].BCformat);
                        }
                    }
                }
                if (Label != null) button_importLabels.Enabled = true;
                generateLabel(-1);
                textBox_templateName.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('\\') + 1);
            }
        }

        private void dataGridView_labels_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            generateLabel(dataGridView_labels.CurrentCell.RowIndex);
        }

        private void generateLabel(int gridLine)
        {
            Bitmap newPicture = new Bitmap(labelWidth, labelHeight, PixelFormat.Format32bppPArgb);
            pictureBox_label.Image = newPicture;
            for (int i = 0; i < Label.Length; i++)
            {
                if (Label[i].type == "text")
                {
                    string fontname = Label[i].fontName;
                    int fontSize = Label[i].fontSize;
                    int posX = Label[i].posX;
                    int posY = Label[i].posY;
                    int rotate = Label[i].rotate;
                    string content = Label[i].content;
                    FontStyle fontStyle = (FontStyle)Label[i].fontStyle;
                    string tmp = "";
                    if (gridLine > -1) tmp = dataGridView_labels.Rows[gridLine].Cells[i - 1].Value.ToString();
                    if (tmp != "") content = tmp;
                    drawText(pictureBox_label.Image, posX, posY, content, fontname, fontSize, rotate, fontStyle);
                }
                else if (Label[i].type == "picture")
                {
                    int posX = Label[i].posX;
                    int posY = Label[i].posY;
                    int rotate = Label[i].rotate;
                    string content = Label[i].content;
                    int width = Label[i].width;
                    int height = Label[i].height;
                    bool transparent = Label[i].transparent;
                    string tmp = "";
                    if (gridLine > -1) tmp = dataGridView_labels.Rows[gridLine].Cells[i - 1].Value.ToString();
                    if (tmp != "") content = tmp;
                    drawPicture(pictureBox_label.Image, posX, posY, content, rotate, width, height, transparent);
                }
                else if (Label[i].type == "barcode")
                {
                    int posX = Label[i].posX;
                    int posY = Label[i].posY;
                    int rotate = Label[i].rotate;
                    string content = Label[i].content;
                    int width = Label[i].width;
                    int height = Label[i].height;
                    BarcodeFormat BCformat = (BarcodeFormat)Label[i].BCformat;
                    string tmp = "";
                    if (gridLine > -1) tmp = dataGridView_labels.Rows[gridLine].Cells[i - 1].Value.ToString();
                    if (tmp != "") content = tmp;
                    drawBarcode(pictureBox_label.Image, posX, posY, width, height, content, BCformat, rotate);
                }
                else if (Label[i].type == "2dcode")
                {
                    int posX = Label[i].posX;
                    int posY = Label[i].posY;
                    int rotate = Label[i].rotate;
                    string content = Label[i].content;
                    int width = Label[i].width;
                    int height = Label[i].height;
                    BarcodeFormat BCformat = (BarcodeFormat)Label[i].BCformat;
                    string tmp = "";
                    if (gridLine > -1) tmp = dataGridView_labels.Rows[gridLine].Cells[i - 1].Value.ToString();
                    if (tmp != "") content = tmp;
                    drawBarcode(pictureBox_label.Image, posX, posY, width, height, content, BCformat, rotate);
                }
            }
        }

        private void button_printCurrent_Click(object sender, EventArgs e)
        {
            if (pictureBox_label.Image == null) return;
            Bitmap bmp = new Bitmap(labelWidth, labelHeight);
            Rectangle rect = new Rectangle(0, 0, labelWidth, labelHeight);
            pictureBox_label.DrawToBitmap(bmp, rect);
            if (!checkBox_toFile.Checked)
            {
                printDialog1 = new PrintDialog();
                printDocument1 = new PrintDocument();
                printDialog1.Document = printDocument1;
                printDocument1.PrintPage += (sender2, args) =>
                {
                    args.Graphics.PageUnit = GraphicsUnit.Pixel;
                    args.Graphics.DrawImage(bmp, rect);
                };
                if (printDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
                else MessageBox.Show("Print Cancelled");
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
                if (printDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
                else MessageBox.Show("Print Cancelled");
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

                if (printDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
                else MessageBox.Show("Print Cancelled");
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
            Bitmap bmp = new Bitmap(labelWidth, labelHeight);
            Rectangle rect = new Rectangle(0, 0, labelWidth, labelHeight);
            pictureBox_label.DrawToBitmap(bmp, rect);
            args.Graphics.PageUnit = GraphicsUnit.Pixel;
            args.Graphics.DrawImage(bmp, rect);
            args.HasMorePages = pagesFrom < pagesTo;
            pagesFrom++;
        }

        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = row.Index.ToString();
            }
            dgv.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void savePage()
        {
            Bitmap bmp = new Bitmap(pictureBox_label.Image.Width, pictureBox_label.Image.Height);
            pictureBox_label.DrawToBitmap(bmp, pictureBox_label.ClientRectangle);
            if (checkBox_toFile.Checked)
            {
                bmp.Save(textBox_saveFileName.Text + dataGridView_labels.CurrentCell.RowIndex.ToString() + ".png", ImageFormat.Png);
            }
        }

        private void checkBox_toFile_CheckedChanged(object sender, EventArgs e)
        {
            textBox_saveFileName.Enabled = checkBox_toFile.Checked;
        }

        private void checkBox_scale_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_scale.Checked) pictureBox_label.SizeMode = PictureBoxSizeMode.CenterImage;
            else pictureBox_label.SizeMode = PictureBoxSizeMode.Zoom;
        }
    }
}
