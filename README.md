# LabelPrint
Prints lots of labels using simple template and list of the data to insert into template
Template and label data are loaded from .CSV file.
Generated pictures are sent to any printer driver or PNG picture.
Printing resolution is set fixed in the template file and must match printer resolution.
Upd: color management added

See readme.odt for more details.

Commandline switches added recently:

/?, /h, /help - print help
/T=file.csv - load template data from file
/L=file.csv - load label data from file
/C - [optional] 1st string of label file is column names (default = no)
/PRN=SystemPrinterName - output to printer (replace spaces with \'_\')
or
/PIC=pictureName - output to pictures
/P=A - print all labels
or
/P=xxx - print label #xxx
or
/P=xxx-yyy - print labels from xxx to yyy
