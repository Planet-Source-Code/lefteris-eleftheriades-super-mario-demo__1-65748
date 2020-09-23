Attribute VB_Name = "mGraphics"
'Functions for drawing gdi graphics
Public Const SRCCOPY = &HCC0020   'Raster operation Replace color values with new ones DestColor = NewColor
Public Const SRCAND = &H8800C6    'Raster operation Logical AND of color values DestColor = DestColor And NewColor
Public Const SRCINVERT = &H660046 'Raster operation Logical XOR of color values DestColor = DestColor XOR NewColor
Public Const Opacity = &H10000    'Applies to AlphaBlend, BLENDFUNCT. Multiply opacity constant with a number[0-255]

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Srchdc As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal BLENDFUNCT As Long) As Long
