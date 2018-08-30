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

	=xxx - print label #xxx

	=xxx-yyy - print labels from #xxx to #yyy


Planned features:
+0) allow both "." and "," float point.
+1) calculate formulas in the numeric input fields
2) show seq.number of the element in the interface (elements list?)
3) allow macros to include other field value (%xxx)
4) allow basic group operations (move left/right/top/down)
5) allow multiselect and edit fields?
6) show coordinates on the mouse cursor on the label picture
7) ??? drag'n'drop objects