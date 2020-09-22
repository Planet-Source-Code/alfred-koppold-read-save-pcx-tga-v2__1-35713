VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Read all pcx and tga-files and save pcx without a dll!"
   ClientHeight    =   4395
   ClientLeft      =   585
   ClientTop       =   1290
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   7275
   WindowState     =   2  'Maximiert
   Begin VB.CheckBox Check1 
      Caption         =   "Compression"
      Height          =   252
      Left            =   9000
      TabIndex        =   15
      Top             =   6360
      Value           =   1  'Aktiviert
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save to TGA"
      Height          =   372
      Left            =   9000
      TabIndex        =   14
      Top             =   5880
      Width           =   1452
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save to BMP"
      Height          =   372
      Left            =   9000
      TabIndex        =   13
      Top             =   4920
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save to PCX"
      Height          =   372
      Left            =   9000
      TabIndex        =   11
      Top             =   5400
      Width           =   1452
   End
   Begin VB.Frame Frame3 
      Caption         =   "Scale"
      Height          =   1092
      Left            =   9000
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
      Begin VB.OptionButton Option7 
         Caption         =   "200%"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "100%"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "50%"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pattern"
      Height          =   975
      Left            =   9000
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
      Begin VB.OptionButton Option8 
         Caption         =   "*.bmp *.gif *.jpg"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "*.pcx"
         Height          =   252
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   852
      End
      Begin VB.OptionButton Option1 
         Caption         =   "*.tga"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1050
      Left            =   9000
      Pattern         =   "*.tga"
      TabIndex        =   3
      Top             =   1560
      Width           =   1932
   End
   Begin VB.DirListBox Dir1 
      Height          =   1170
      Left            =   9000
      TabIndex        =   2
      Top             =   480
      Width           =   1932
   End
   Begin VB.DriveListBox Drive1 
      Height          =   360
      Left            =   9000
      TabIndex        =   1
      Top             =   120
      Width           =   1932
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   2172
      Left            =   240
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   100
      Width           =   1932
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
Dim b As New SaveTGA
Dim Weg As String
Dim a As New OpenSaveDLG
Dim Kompr As Boolean
On Error Resume Next
Kompr = Check1.Value
a.DialogTitle = "Save-Filename"
a.Filter = "TGA Dateien(*.tga)" & Chr$(0) & "*.tga" & Chr$(0)
a.Show 1
If a.ReturnPath = "" Then Exit Sub
If LCase(Right(a.ReturnPath, 3)) = ".tga" Then
Weg = a.ReturnPath
Else
Weg = a.ReturnPath & ".tga"
Set a = Nothing
End If
b.SaveTGAinFile Weg, Pic1, Kompr
Set b = Nothing
End Sub

Private Sub Command2_Click()
Dim Kompr As Boolean
Dim b As New SavePCX
Dim Weg As String
Dim a As New OpenSaveDLG

On Error Resume Next
Kompr = Check1.Value
a.DialogTitle = "Save-Filename"
a.Filter = "PCX Dateien(*.pcx)" & Chr$(0) & "*.pcx" & Chr$(0)
a.Show 1
If a.ReturnPath = "" Then Exit Sub
If LCase(Right(a.ReturnPath, 3)) = ".pcx" Then
Weg = a.ReturnPath
Else
Weg = a.ReturnPath & ".pcx"
Set a = Nothing
End If
b.SavePCXinFile Weg, Pic1, Kompr
Set b = Nothing
End Sub

Private Sub Command3_Click()
Dim Weg As String
Dim a As New OpenSaveDLG
On Error Resume Next
a.Filter = "BMP Dateien(*.bmp)" & Chr$(0) & "*.bmp" & Chr$(0)
a.DialogTitle = "Save-Filename"
a.Show 1
If a.ReturnPath = "" Then Exit Sub
If LCase(Right(a.ReturnPath, 3)) = ".bmp" Then
Weg = a.ReturnPath
Else
Weg = a.ReturnPath & ".bmp"
Set a = Nothing
End If
SavePicture Pic1, Weg
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim Pfad As String
If Right(File1.Path, 1) <> "\" Then
Pfad = File1.Path & "\" & File1.FileName
Else
Pfad = File1.Path & File1.FileName
End If
Dim Scaling As Single
If Option5.Value = True Then Scaling = 2 ' 50%
If Option6.Value = True Then Scaling = 1 '100%
If Option7.Value = True Then Scaling = 0.5 '200%
Pic1.Cls

If Option1.Value = True Then 'TGA-File
Dim tgaFile As New LoadTGA
If Option6.Value = False Then
tgaFile.Autoscale = False
Else
tgaFile.Autoscale = True
End If

tgaFile.LoadTGA Pfad
If tgaFile.IsTGA = True Then 'is it a TGA-File?
Form1.Pic1.Width = tgaFile.TGAWidth / Scaling
Form1.Pic1.Height = tgaFile.TGAHeight / Scaling
tgaFile.DrawTGA Form1.Pic1
End If
ElseIf Option2.Value = True Then 'PCX-File
Dim pcxFile As New LoadPCX

If Option6.Value = False Then
pcxFile.Autoscale = False
Else
pcxFile.Autoscale = True
End If
pcxFile.LoadPCX Pfad
If pcxFile.IsPCX = True Then 'is it a PCX-File?
pcxFile.ScaleMode = 1
Form1.Pic1.Width = pcxFile.PCXWidth / Scaling
Form1.Pic1.Height = pcxFile.PCXHeight / Scaling
pcxFile.DrawPCX Form1.Pic1
End If
ElseIf Option8.Value = True Then
Pic1.Picture = LoadPicture(Pfad)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Option1_Click()
File1.Pattern = "*.tga"
End Sub

Private Sub Option2_Click()
File1.Pattern = "*.pcx"
End Sub


Private Sub Option8_Click()
File1.Pattern = "*.bmp*;*.gif;*.jpg"
End Sub
