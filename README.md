# LabelPrint
Prints lots of labels using simple template and list of the data to insert into template
Template and label data are loaded from .CSV file.
Generated pictures are sent to any printer driver or PNG picture.
Upd: color management added

See readme.odt for more details.

Commandline switches added recently:

/?, /h, /help - print help\r\n
/T=file.csv - load template data from file\r\n
/L=file.csv - load label data from file\r\n
/C - [optional] 1st string of label file is column names (default = no)\r\n
/PRN=SystemPrinterName - output to printer (replace spaces with \'_\')\r\n
or
/PIC=pictureName - output to pictures\r\n
/P=A - print all labels\r\n
or
/P=xxx - print label #xxx\r\n
or
/P=xxx-yyy - print labels from xxx to yyy