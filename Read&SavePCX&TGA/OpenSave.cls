VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenSaveDLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Für Datei öffnen, Datei speichern
Private OwnerVAR As Long
Private FilterVAR As String
Private DialogTitleVAR As String
Private InitDirVAR As String
Private ReturnPathVAR As String
Private ShortFileNameVAR As Boolean
Private FileNameVAR As String
Private ReadOnlyVAR As Boolean
Private OldStyleVAR As Boolean
Private CancelErrorVAR As Boolean

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Private Const OFN_READONLY = &H1
Private Const OFN_EXPLORER = &H80000
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const VER_PLATFORM_WIN32_NT = 2

                         
Private Declare Function GetOpenFileName _
                         Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
                         (pOpenfilename As OPENFILENAME) As Long


Private Declare Function GetSaveFileName _
                         Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
                         (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetShortPathName _
                        Lib "kernel32" Alias "GetShortPathNameA" _
                        (ByVal lpszLongPath As String, _
                        ByVal lpszShortPath As String, _
                        ByVal cchBuffer As Long) As Long


'VersionsInfos über das BS!
Private Declare Function GetVersionEx _
                        Lib "kernel32" Alias "GetVersionExA" _
                        (lpVersionInformation As OSVERSIONINFO) As Long


Public Property Get ShortFileName() As Boolean
 ShortFileName = ShortFileNameVAR
End Property



Public Property Get OldStyle() As Boolean
 OldStyle = OldStyleVAR
End Property



Public Property Let Filter(Wert As String)
 FilterVAR = Trim$(Wert)
End Property

Public Property Let InitDir(Wert As String)
 InitDirVAR = Trim$(Wert)
End Property
Public Property Let ShortFileName(Wert As Boolean)
 ShortFileNameVAR = Wert
End Property

Public Property Let ReturnPath(Wert As String)
 ReturnPathVAR = Trim$(Wert)
End Property

Public Property Let FileName(Wert As String)
 FileNameVAR = Trim$(Wert)
End Property


Public Property Let Owner(Wert As Long)
 OwnerVAR = Wert
End Property



Public Property Let ReadOnly(Wert As Boolean)
 ReadOnlyVAR = Wert
End Property


Public Property Let OldStyle(Wert As Boolean)
 OldStyleVAR = Wert
End Property

Public Property Let DialogTitle(Wert As String)
 DialogTitleVAR = Trim$(Wert)
End Property

Public Property Get CancelError() As Boolean
 CancelError = CancelErrorVAR
End Property

Public Property Let CancelError(Wert As Boolean)
 CancelErrorVAR = Wert
End Property




Public Property Get Filter() As String
 Filter = FilterVAR
End Property

 

Public Property Get DialogTitle() As String
 DialogTitle = DialogTitleVAR
End Property
Public Property Get InitDir() As String
 InitDir = InitDirVAR
End Property
Public Property Get FileName() As String
 FileName = FileNameVAR
End Property
Public Property Get ReturnPath() As String
If ShortFileName Then ' Es wird das Format 8:3 gewünscht.
 ReturnPath = GetShortPathNameFunc(ReturnPathVAR)
Else ' Einfachste Übergabe.
 ReturnPath = Trim$(ReturnPathVAR)
End If
End Property
Public Property Get Owner() As Long
 Owner = OwnerVAR
End Property

Function Show(Variante As Long) As Boolean
Dim DateiÖffnenAPIRuf As OPENFILENAME

With DateiÖffnenAPIRuf
.lStructSize = Len(DateiÖffnenAPIRuf)
.hWnd = Owner
.lpstrTitle = DialogTitle
.lpstrFile = FileName & String$(250 - Len(FileName), 0)
.nMaxFile = 255
.lpstrFileTitle = String$(255, 0)
.nMaxFileTitle = 255
.lpstrFilter = Filter
.nFilterIndex = 1
.lpstrInitialDir = InitDir
.Flags = OFN_ALLOWMULTISELECT

 If ReadOnlyVAR Then .Flags = OFN_READONLY
 If OldStyle Then
  .Flags = .Flags Or OFN_ALLOWMULTISELECT
 Else
  .Flags = .Flags Or OFN_EXPLORER
 End If
End With


Dim Rückgabe As Long

If Variante = 0 Then
 Rückgabe = GetOpenFileName(DateiÖffnenAPIRuf)
 Else
  Rückgabe = GetSaveFileName(DateiÖffnenAPIRuf)

End If

Select Case Rückgabe
Case 1 'es wurde ein Wert übergeben
 Show = True
 ReturnPath = Left(DateiÖffnenAPIRuf.lpstrFile, InStr(DateiÖffnenAPIRuf.lpstrFile, Chr(0)) - 1)

Case 0 ' Der Benutzer hat auf Abbrechen geklickt
 ReturnPath = vbNullString
 If CancelError Then
  Err.Clear
  Err.Raise 18
  Show = False
 End If
 
Case Else
 Show = False
 ReturnPath = vbNullString ' Das ist nicht gut,sollte nie passieren :-)
End Select
End Function

Private Function ByteToString(BypteAR() As Byte) As String
Dim BCount As Long
Dim ByteVal As Long
Dim OutPutVAR As String
    
BCount = LBound(BypteAR)
    
While BCount <= UBound(BypteAR)
 ByteVal = BypteAR(BCount)
 If ByteVal = 0 Then
  ByteToString = OutPutVAR
  Exit Function
 Else
  OutPutVAR = OutPutVAR & Chr$(ByteVal)
 End If
 BCount = BCount + 1
Wend
    
ByteToString = OutPutVAR
End Function

Private Function IsWinNT() As Boolean
Dim LokalesOS As OSVERSIONINFO

LokalesOS.dwOSVersionInfoSize = Len(LokalesOS)
Call GetVersionEx(LokalesOS)

If LokalesOS.dwPlatformId = VER_PLATFORM_WIN32_NT Then
 IsWinNT = True
Else
 IsWinNT = False
End If

End Function

Private Function GetShortPathNameFunc(DateiName As String) As String
Dim sBuffer As String, lLen As Long
sBuffer = Space$(1024)
lLen = GetShortPathName(DateiName, sBuffer, Len(sBuffer))
GetShortPathNameFunc = Left$(sBuffer, lLen)
End Function


