VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SavePCX"
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
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO8Bit, ByVal wUsage As Long) As Long

Private Bildhöhe As Long
Private Bildbreite As Long
Private Bildbreiteright As Long
Private CompPCX() As Byte
Private Compressed As Boolean
Private Bitcbr As Long 'BPL

Private Const DIB_RGB_COLORS = 0
Private Const BI_RGB = 0&

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
Private Type BITMAPINFO8Bit
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type
Private Type PCXHeader
    ZSoftID As Byte
    Version As Byte
    RLE As Byte
    BitsPerPixel As Byte
    XMIN As Integer
    YMIN As Integer
    XMAX As Integer
    YMAX As Integer
    HRes As Integer
    VRes As Integer
    Palette16(0 To 47) As Byte
    Reserved As Byte
    nPlanes As Byte
    BytesPerLine As Integer
    PaletteType As Integer
    info(0 To 57) As Byte
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type



Public Sub SavePCXinFile(FileName As String, Pic As PictureBox, Optional Compress As Boolean = True)
    Dim a As Long
    Dim i As Long
    Dim iBitmap As Long, iDC As Long
    Dim b As Long
    Dim TempPal(0 To 255) As RGBQUAD
    Dim bBits() As Byte
    Dim Dibbreite As Long
    Dim pcxbreite As Long
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
    If (Bildbreiteright * 3) Mod 2 = 0 Then
    pcxbreite = Bildbreiteright * 3
    Else
    pcxbreite = Bildbreiteright * 3 + 1
    End If
    
    If Bildbreiteright Mod 2 = 0 Then
    Bitcbr = Bildbreiteright
    Else
    Bitcbr = Bildbreiteright + 1
    End If
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

    SavePCX FileName, Übergabe(), Bitcbr, Bildhöhe

End Sub
Private Sub SavePCX(DSTFile As String, BitmapBytes() As Byte, PCXWidth As Long, PCXHeight As Long)
    Dim bBytes As Long
    Dim PCXMap() As Byte
    Dim FileHeader As PCXHeader
    Dim i As Long
    Dim PCXRed() As Byte
    Dim PCXGreen() As Byte
    Dim PCXBlue() As Byte
    Dim Übergabe() As Byte
    Dim a As Long
    
    FileHeader.ZSoftID = 10
    FileHeader.Version = 5
    Select Case Compressed
    Case False
    FileHeader.RLE = 0
    Case True
    FileHeader.RLE = 1
    End Select
    FileHeader.BitsPerPixel = 8
    FileHeader.nPlanes = 3
    FileHeader.XMIN = 1
    FileHeader.YMIN = 1
    FileHeader.XMAX = Bildbreiteright
    FileHeader.YMAX = PCXHeight
    FileHeader.HRes = 96
    FileHeader.VRes = 96
    FileHeader.Reserved = 0
    FileHeader.PaletteType = 0
    
            FileHeader.BytesPerLine = PCXWidth
    
            MakePCX24 BitmapBytes, PCXMap, PCXWidth, PCXHeight, PCXWidth
            If Dir(DSTFile) <> "" Then Kill DSTFile
            If Compressed = True Then
            CompressPCX PCXMap, PCXWidth
            End If
            'For i = 1 To 100
            'Debug.Print PCXMap(i)
            'Next i
            Open DSTFile For Binary As #1
                Put #1, , FileHeader
            If Compressed = False Then
                Put #1, , PCXMap
            Else
                Put #1, , CompPCX
            End If
            Close
End Sub

Private Sub MakePCX24(Feld() As Byte, PCXMap() As Byte, nWidth As Long, nHeight As Long, BytesPerLine As Long)
    Dim x As Long, y As Long, nStartPos As Long
    Dim High As Long
        Dim quad() As RGBTRE
        ReDim quad(nWidth * nHeight)
        CopyMemory quad(0), Feld(0), (nWidth * nHeight * 3)
        ReDim PCXMap(nWidth * nHeight * 3 - 1)
    For y = 0 To nHeight - 1
        For x = 0 To nWidth - 1
        High = nHeight - y - 1
            nStartPos = High * BytesPerLine * 3 + x '0-31 96-127 192-
            With quad((x + (nHeight - 1 - y) * nWidth)) ' zuerst 992- 1023 Ende 0-31
                 PCXMap(nStartPos) = .rgbRed
                PCXMap(nStartPos + BytesPerLine) = .rgbGreen
                PCXMap(nStartPos + BytesPerLine * 2) = .rgbBlue
            End With
        
        Next x
    Next y
    
End Sub
Private Sub CompressPCX(ArraytoComp() As Byte, Linewidth As Long)
Dim i As Long
Dim b As Long
Dim Zähler As Long
Dim Stand As Long
Dim wo As Long
Dim Linezahl As Long

ReDim CompPCX((UBound(ArraytoComp) * 2) + ((UBound(ArraytoComp) \ Linewidth) * 2)) ' größtmögliche Größe
'Linezahl = 1
Do While i <= UBound(ArraytoComp)
If i + 1 <= UBound(ArraytoComp) Then
Do Until ArraytoComp(i) <> ArraytoComp(i + 1) Or Zähler = 62
If i + 1 = (Linewidth * Linezahl) And Linewidth Mod 2 <> 0 Then Exit Do ' +1 da Grund = 0
i = i + 1
Zähler = Zähler + 1
If i = UBound(ArraytoComp) Then Exit Do
Loop
Else

End If
If Zähler = 0 Then

If ArraytoComp(i) < 192 Then
CompPCX(b) = ArraytoComp(i) ' 1* kleiner 192
wo = b
b = b + 1
If i + 1 = (Linewidth * Linezahl) And Linewidth Mod 2 <> 0 Then 'falls Zeilenende und ungerade Zahl
CompPCX(b) = 0
wo = b
b = b + 1
Linezahl = Linezahl + 1
End If
Else
CompPCX(b) = 193
CompPCX(b + 1) = ArraytoComp(i)  '1* größer 192
wo = b + 1
b = b + 2
If i + 1 = (Linewidth * Linezahl) And Linewidth Mod 2 <> 0 Then 'falls Zeilenende und ungerade Zahl
CompPCX(b) = 0
wo = b
b = b + 1
Linezahl = Linezahl + 1
End If
End If

Else
CompPCX(b) = 192 + Zähler + 1 ' bis 255 ist Zähler = 62
CompPCX(b + 1) = ArraytoComp(i) 'Zähler * Byte
wo = b + 1
b = b + 2
If i + 1 = (Linewidth * Linezahl) And Linewidth Mod 2 <> 0 Then 'falls Zeilenende und ungerade Zahl
CompPCX(b) = 0
wo = b
b = b + 1
Linezahl = Linezahl + 1
End If
End If

Zähler = 0 'Zähler zurücksetzen
i = i + 1
Loop
ReDim Preserve CompPCX(wo) 'zuschneiden
ReDim ArraytoComp(wo)
CopyMemory ArraytoComp(0), CompPCX(0), wo + 1
End Sub
Private Sub GetIconBits(bBits() As Byte, nDC As Long, nBitmap As Long, CopyArr() As RGBQUAD, Height As Long, Breite As Long)
    Dim BI As BITMAPINFO8Bit
    Dim Test As Long
        BI.bmiHeader.biHeight = 0 - Height
    BI.bmiHeader.biBitCount = 24
    BI.bmiHeader.biCompression = BI_RGB
    BI.bmiHeader.biPlanes = 1
    BI.bmiHeader.biWidth = Breite
    BI.bmiHeader.biSize = Len(BI.bmiHeader)
    Test = GetDIBits(nDC, nBitmap, 0, Height, bBits(0), BI, DIB_RGB_COLORS)
End Sub

