Label data:

1st string: Column headers
2nd string: 1st label data
3rd string: 2nd label data

====label data file====
String-1;Picture-1;Barcode-1;2Dcode-1;
string1;file1.png;123456789111;987654321111;
string2;file2.png;123456789222;987654321222;
====end====


Template data:

1st string: label definition and size
2nd string: 1st object definition and parameters
3rd string: 2nd object definition and parameters

====label template file====
  label; width; height;
   text; posX;  posY; [rotate]; default_text; fontName; fontSize; [fontStyle];
picture; posX;  posY; [rotate]; default_file; [width];  [height]; [transparent];
barcode; posX;  posY; [rotate]; default_data; width;    height;   bcFormat; [additional_features]
====end====

Parameters:

[] - can be set as "0" to use default value
object content data will be taken from template "default_" field if label file has empty field for it.

fontStyle:
	0 = REGULAR
	1 = BOLD
	2 = ITALIC
	4 = UNDERLINE
	8 = STRIKEOUT

transparent:
	0 = no conversion
	1 = convert white to transparent

bcFormat:
1D
	2 = CODABAR
	4 = CODE39
	8 = CODE93
	16 = CODE128
	64 = EAN8
	128 = EAN13
	256 = ITF
	4096 = RSS_14
	8192 = RSS_EXPANDED
	16384 = UPC_A
	32768 = UPC_E
	65535 = UPC_EAN_EXTENSION
	131072 = MSI
	262144 = PLESSEY
	524288 = IMB
2D
	1 = AZTEC
	32 = DATAMATRIX
	512 = MAXICODE
	1024 = PDF417
	2048 = QRCODE

additional_features:
AZTEC_LAYERS
	[-4, 32]
	Specifies the required number of layers for an Aztec code. A negative 
	number (-1, -2, -3, -4) specifies a compact Aztec code. 0 indicates 
	to use the minimum number of layers (the default). A positive number 
	(1, 2, .. 32) specifies a normal (non-compact) Aztec code. (Type 
	Integer, or String representation of the integer value).
ERROR_CORRECTION
	[0, 8]
	Specifies what degree of error correction to use, for example in 
	QR Codes. Type depends on the encoder. For example for QR codes it's 
	type ErrorCorrectionLevel. For Aztec it is of type Integer, 
	representing the minimal percentage of error correction words. 
	For PDF417 it is of type Integer, valid values being 0 to 8. In all 
	cases, it can also be a String representation of the desired value as 
	well. Note: an Aztec symbol should have a minimum of 25% EC words.
MARGIN
	[int]
	Specifies margin, in pixels, to use when generating the barcode. 
	The meaning can vary by format; for example it controls margin before 
	and after the barcode horizontally for most 1D formats. (Type Integer, 
	or String representation of the integer value).
PDF417_ASPECT_RATIO
	[1, 4]
	Aspect ratio to use. Default = 4.
QR_VERSION
	[1, 40] ??
	Specifies the exact version of QR code to be encoded. (Type Integer, 
	or String representation of the integer value).
CHARACTER_SET
	[string]
	Specifies what character encoding to use where applicable (type String)
PDF417_COMPACTION
	[string]
	Specifies what compaction mode to use for PDF417 (type Compaction or 
	String value of one of its enum values).
CODE128_FORCE_CODESET_B
	[true, false]
	If true, don't switch to codeset C for numbers
DISABLE_ECI
	[true, false]
	Don't append ECI segment.
GS1_FORMAT
	[true, false]
	Specifies whether the data should be encoded to the GS1 standard 
	(type Boolean, or "true" or "false" String value).
PDF417_COMPACT
	[true, false]
	Specifies whether to use compact mode for PDF417 (type Boolean, or 
	"true" or "false" String value).
DATA_MATRIX_SHAPE
	[?]
	Specifies the matrix shape for Data Matrix (type SymbolShapeHint)
PDF417_DIMENSIONS
	[?]
	Specifies the minimum and maximum number of rows and columns for 
	PDF417 (type Dimensions).
DATA_MATRIX_DEFAULT_ENCODATION
	[?]
	Encodation for DataMatrix. Default = ASCII.
