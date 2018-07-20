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

1st string: label mark and size
2nd string: 1st object mark and parameters
3rd string: 2nd object mark and parameters

====label template file====
  label; width; height;
   text; posX;  posY; [rotate]; default_text; fontName; fontSize; [fontStyle];
picture; posX;  posY; [rotate]; default_file; [width];  [height]; [transparent];
barcode; posX;  posY; [rotate]; default_data; width;    height;   bcFormat;
 2Dcode; posX;  posY; [rotate]; default_data; width;    height;   bcFormat;
====end====

Parameters:

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
1 = AZTEC
2 = CODABAR
4 = CODE39
8 = CODE93
16 = CODE128
32 = DATAMATRIX
64 = EAN8
128 = EAN13
256 = ITF
512 = MAXICODE
1024 = PDF417
2048 = QRCODE
4096 = RSS_14
8192 = RSS_EXPANDED
16384 = UPC_A
32768 = UPC_E
65535 = UPC_EAN_EXTENSION
131072 = MSI
262144 = PLESSEY
524288 = IMB
