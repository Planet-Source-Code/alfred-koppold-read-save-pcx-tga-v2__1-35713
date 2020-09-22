VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaveTGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Autor: ALKO
'e-mail: alfred.koppold@freenet.de
Option Explicit
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Const DIB_RGB_COLORS = 0
Private Const BI_RGB = 0&
Private Compressed As Boolean
Private Bildhöhe As Long
Private Bitcbr As Long
Private Bildbreite As Long
Private Bildbreiteright As Long
Private Type tgaheader                  'declare the header
  info       As Byte      'length of image information block
  colortyp   As Byte      'DAC table or BGR format
  imagetyp   As Byte      'compressed or uncompressed
  origin     As Integer         'first entry in the DAC table
  colnumber  As Integer         'number of colors in the DAC table
  entrybits  As Byte      'entry size in the DAC table
  xvalue     As Integer         'x co-ordinate lower left corner
  yvalue     As Integer         'y co-ordinate lower left corner
  widt       As Integer         'image width
  Height     As Integer         'image height
  pixelsize  As Byte      'number of bits per pixel
  descriptor As Byte      'image descriptor
End Type
Private CompTGA() As Byte
Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Private Type RGBTRE
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
End Type
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Private Type BITMAPINFO8Bit
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type
Public Sub SaveTGAinFile(FileName As String, Pic As PictureBox, Optional Compress As Boolean = True)
    Dim a As Long
    Dim i As Long
    Dim iBitmap As Long, iDC As Long
    Dim b As Long
    Dim TempPal(0 To 255) As RGBQUAD
    Dim bBits() As Byte
    Dim Dibbreite As Long
    Dim TGAbreite As Long
        Dim Groesse As Long
    Dim Style As Long
      Dim aDC As Long  'source bitmap (color)
  Dim aPrevBmp As Long  'Holds previous bitmap in source DC
    Dim bi24BitInfo As BITMAPINFO
Dim Übergabe() As Byte
Compressed = Compress
    a = Pic.Parent.ScaleMode
    b = Pic.ScaleMode
    Pic.Parent.ScaleMode = 3
    Pic.ScaleMode = 3
    Bildhöhe = Pic.ScaleHeight
    Bildbreite = Pic.ScaleWidth
    Pic.Parent.ScaleMode = a
    Pic.ScaleMode = b
    Bildbreiteright = Bildbreite

            ' 24 BPP
    
    If Bildbreite * 3 Mod 4 = 0 Then  '24 bit
    Dibbreite = Bildbreite * 3
    Else
    Dibbreite = Bildbreite * 3
    Dibbreite = Dibbreite \ 4
    Dibbreite = Dibbreite + 1
    Dibbreite = Dibbreite * 4

    End If
    TGAbreite = Bildbreiteright * 3
    Bitcbr = Bildbreiteright
        ReDim bBits((Dibbreite) * Bildhöhe - 1)

        With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = Bildbreiteright
        .biHeight = Bildhöhe
    End With

    iDC = CreateCompatibleDC(Pic.hdc)
    iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject iDC, iBitmap
    BitBlt iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, Pic.hdc, 0, 0, vbSrcCopy

    GetIconBits bBits(), iDC, iBitmap, TempPal(), Bildhöhe, Bildbreiteright
    ReDim Übergabe(Bitcbr * Bildhöhe * 3 - 1)
  If Bildbreiteright * 3 <> Dibbreite Then
  For i = 0 To Bildhöhe - 1
  CopyMemory Übergabe(i * Bitcbr * 3), bBits(i * Dibbreite), Bitcbr * 3
  Next i

  Else
    CopyMemory Übergabe(0), bBits(0), UBound(bBits) + 1
  End If
        DeleteDC iDC
    DeleteObject iBitmap

    SaveTGA FileName, Übergabe(), Bitcbr, Bildhöhe

End Sub

Private Sub MakeTGA24(Feld() As Byte, TGAMap() As Byte, nWidth As Long, nHeight As Long, BytesPerLine As Long)
ReDim TGAMap(UBound(Feld))
CopyMemory TGAMap(0), Feld(0), UBound(Feld) + 1

End Sub

Private Sub GetIconBits(bBits() As Byte, nDC As Long, nBitmap As Long, CopyArr() As RGBQUAD, Height As Long, Breite As Long)
    Dim BI As BITMAPINFO8Bit
    Dim Test As Long
        BI.bmiHeader.biHeight = Height
    BI.bmiHeader.biBitCount = 24
    BI.bmiHeader.biCompression = BI_RGB
    BI.bmiHeader.biPlanes = 1
    BI.bmiHeader.biWidth = Breite
    BI.bmiHeader.biSize = Len(BI.bmiHeader)
    Test = GetDIBits(nDC, nBitmap, 0, Height, bBits(0), BI, DIB_RGB_COLORS)
End Sub
Private Sub SaveTGA(DSTFile As String, BitmapBytes() As Byte, TGAWidth As Long, TGAHeight As Long)
    Dim bBytes As Long
    Dim TGAMap() As Byte
    Dim FileHeader As tgaheader
    Dim i As Long
    Dim TGARed() As Byte
    Dim TGAGreen() As Byte
    Dim TGABlue() As Byte
    Dim Übergabe() As Byte
    Dim a As Long
    
    FileHeader.colortyp = 0
    Select Case Compressed
    Case False
    FileHeader.imagetyp = 2 'ohne Komprimerung
    Case True
    FileHeader.imagetyp = 10 'mit Komprimerung
    End Select
    FileHeader.entrybits = 0
    FileHeader.xvalue = 0
    FileHeader.yvalue = 0
    FileHeader.widt = Bildbreiteright
    FileHeader.Height = TGAHeight
    FileHeader.pixelsize = 24
            MakeTGA24 BitmapBytes, TGAMap, TGAWidth, TGAHeight, TGAWidth
            If Dir(DSTFile) <> "" Then Kill DSTFile
            If Compressed = True Then
            CompressTGA TGAMap, TGAWidth
            End If
            'For i = 1 To 100
            'Debug.Print TGAMap(i)
            'Next i
            Open DSTFile For Binary As #1
                Put #1, , FileHeader
            If Compressed = False Then
                Put #1, , TGAMap
            Else

                Put #1, , TGAMap
            End If
            Close
End Sub

Private Sub CompressTGA(ArraytoComp() As Byte, Linewidth As Long)
Dim i As Long
Dim RGBArray() As RGBTRE
Dim b As Long
Dim Zähler As Long
Dim ungleich As Long
Dim Stand As Long
Dim wo As Long
Dim Linezahl As Long
Dim ia As Long
Dim Zählera As Long
Dim CompTGA() As Byte
Dim Dateinummer As Long
ReDim RGBArray((UBound(ArraytoComp) + 1) / 3)
CopyMemory RGBArray(0).rgbBlue, ArraytoComp(0), UBound(ArraytoComp) + 1

ReDim CompTGA((UBound(ArraytoComp) * 2) + ((UBound(ArraytoComp) \ Linewidth) * 2)) ' größtmögliche Größe
Linezahl = 1
Do While i <= UBound(RGBArray)
Zähler = 0
Zählera = 0
If i + 1 <= UBound(RGBArray) Then
Do Until RGBArray(i).rgbBlue <> RGBArray(i + 1).rgbBlue Or RGBArray(i).rgbGreen <> RGBArray(i + 1).rgbGreen Or RGBArray(i).rgbRed <> RGBArray(i + 1).rgbRed Or Zähler = 127
If i + 1 = (Linewidth * Linezahl) Then Exit Do  ' +1 da Grund = 0
i = i + 1
Zähler = Zähler + 1
If i = UBound(RGBArray) Then Exit Do
Loop
Else
'Make Nothing
End If
If Zähler = 0 Then
ia = i
If (ia + 1) <= CLng(UBound(RGBArray)) Then
Do While RGBArray(ia).rgbBlue <> RGBArray(ia + 1).rgbBlue Or RGBArray(ia).rgbGreen <> RGBArray(ia + 1).rgbGreen Or RGBArray(ia).rgbRed <> RGBArray(ia + 1).rgbRed
ia = ia + 1
Zählera = Zählera + 1
If ia + 1 = Linewidth * Linezahl Then Exit Do
If ia = UBound(RGBArray) Then Exit Do
Loop
End If
If Zählera > 0 Then
CompTGA(b) = Zählera - 1
CopyMemory CompTGA(b + 1), RGBArray(i).rgbBlue, (Zählera) * 3
wo = Zählera * 3
b = b + 1 + wo
i = i + Zählera - 1
End If
If i + 1 = (Linewidth * Linezahl) Then 'falls Zeilenende
Linezahl = Linezahl + 1
If Linezahl > (UBound(RGBArray) + 1) / Linewidth Then Exit Do
End If

Else
CompTGA(b) = 127 + Zähler + 1 ' bis 255 ist Zähler = 127
CopyMemory CompTGA(b + 1), RGBArray(i).rgbBlue, 3
'CompTGA(b + 1) = ArraytoComp(i) 'Zähler * Byte
b = b + 4
If i + 1 = (Linewidth * Linezahl) Then  'falls Zeilenende
Linezahl = Linezahl + 1
If Linezahl > (UBound(ArraytoComp) + 1) / Linewidth Then Exit Do
End If
End If

Zähler = 0 'Zähler zurücksetzen
i = i + 1
Loop
ReDim Preserve CompTGA(b) 'zuschneiden
'For i = 0 To UBound(CompTGA)
'Form1.Text1 = Form1.Text1 & " " & CompTGA(i)
'Next i
ReDim ArraytoComp(b - 1)
CopyMemory ArraytoComp(0), CompTGA(0), b
End Sub


