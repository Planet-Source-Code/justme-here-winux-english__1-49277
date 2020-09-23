VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form TXPreader 
   BorderStyle     =   0  'None
   Caption         =   "TXP Reader - Media Player"
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   ClipControls    =   0   'False
   Icon            =   "TXPreader.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox setc 
      Height          =   645
      Left            =   5520
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C15D1A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   9480
      MouseIcon       =   "TXPreader.frx":0CCA
      MousePointer    =   99  'Custom
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   10
      Top             =   8160
      Width           =   135
   End
   Begin RichTextLib.RichTextBox Rt 
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"TXPreader.frx":0FD4
   End
   Begin VB.ListBox info 
      Height          =   840
      Left            =   6720
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox pass 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pLine 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   840
      Picture         =   "TXPreader.frx":1042
      ScaleHeight     =   90
      ScaleWidth      =   24000
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   24000
   End
   Begin RichTextLib.RichTextBox Ligne 
      Height          =   1695
      Left            =   1800
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2990
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"TXPreader.frx":1D68
   End
   Begin VB.PictureBox pictemp 
      Height          =   1335
      Left            =   480
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox txttemp 
      Height          =   1455
      Left            =   2280
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"TXPreader.frx":1DD6
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFBE71&
      Height          =   3325
      Left            =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C15D1A&
      BorderWidth     =   4
      Height          =   3060
      Left            =   30
      Top             =   240
      Width           =   4290
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   200
      Left            =   9480
      TabIndex        =   13
      ToolTipText     =   "Fermer"
      Top             =   40
      Width           =   200
   End
   Begin VB.Image Cur 
      Height          =   480
      Left            =   6960
      Picture         =   "TXPreader.frx":1E44
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   360
      Width           =   15
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   120
      Top             =   360
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Piccount 
      Caption         =   "0"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MediaPlayerCtl.MediaPlayer Son 
      Height          =   30
      Left            =   6240
      TabIndex        =   4
      Top             =   60000
      Visible         =   0   'False
      Width           =   30
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C15D1A&
      Caption         =   "TXP Reader"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      ToolTipText     =   "To enlarge, double click"
      Top             =   20
      Width           =   9715
   End
End
Attribute VB_Name = "TXPreader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Winux Graphic User Interface for Windows based systems
'Copyright (C) 2002-2003 Winux Team
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'A copy of this licence is available in root\system directory.
'http://www.winux.free.fr or tex_winux@hotmail.com for more details.

'le fichier à lire
Public txpfile

Private Const BASE = 65521
Private Const CHUNK_SIZE = 2048
Private Const NMAX = 5552
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const MAX_PATH = 256&
Private Const REG_SZ = 1
Dim PAKFile As String
Dim FileListStart As Long
Private s(0 To 255) As Long
Private i As Long
Private j As Long
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const EM_UNDO = &HC7
Private Const EM_CANUNDO = &HC6
Private Const WM_COPY& = &H301
Private Const WM_CUT& = &H300
Private Const WM_PASTE& = &H302
Dim FFilename As String
Dim imgText(999) As String
Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1
Dim lignet As String
Dim imouse As Integer
Private Type Size
    cx As Long
    cy As Long
End Type
Dim htxt As String

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MM_ANISOTROPIC = 8
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Public Function GetLongName(ByVal sShortFileName As String) As String
    Dim lRetVal As Long
    Dim sLongFileName As String
    Dim iLen As Integer
    sLongFileName = SPACE(255)
    iLen = Len(sLongFileName)
    lRetVal = GetLongPathName(sShortFileName, sLongFileName, iLen)
    GetLongName = Left(sLongFileName, lRetVal)
End Function
Private Function getTempName(Optional anExt As String = "tmp") As String
    Dim tempPath    As String
    Dim filename    As String
    Dim i           As Long
    
    Const validChars As String = "123567890qwertyuiopasdfghjklzxcvbnm"
    tempPath = String$(255, " ")
    GetTempPath 255, tempPath
    tempPath = Left$(tempPath, InStr(tempPath, Chr$(0)) - 1)
    filename = SPACE(12)
    Mid$(filename, 1, 1) = "T"
    Mid$(filename, Len(filename) - Len(anExt), 1) = "."
    Mid$(filename, Len(filename) - Len(anExt) + 1, Len(anExt)) = anExt
    Randomize
    For i = 2 To Len(filename) - 4
        Mid$(filename, i, 1) = Mid$(validChars, CLng(Rnd() * (Len(validChars)) + 1), 1)
    Next i
    tempPath = tempPath & filename
    getTempName = tempPath
End Function
Public Function FindOppAsc(Value As Integer) As Integer
    If Value <> 128 Then
        FindOppAsc = 255 - Value
    Else
        FindOppAsc = Value
    End If
End Function

Public Function Converter(xString As String) As String
    On Error GoTo FinaliseError
    For cCode = 1 To Len(xString)
        conv = conv + (100 / Len(xString))
        Converter = Converter + Chr(FindOppAsc(Asc(Mid(xString, CInt(cCode), 1))))
    Next cCode
    Exit Function
FinaliseError:
    MsgBox "Erreur de cryptage / décryptage"
End Function

Private Sub Huffman_Progress(Procent As Integer)
On Error Resume Next
  Shape2.Width = Shape1.Width * Procent / 100
  DoEvents
End Sub

Private Sub Form_Load()
  Set Huffman = New clsHuffman
    Clipboard.Clear
    Clipboard.SetData pLine.Picture
    SendMessage Ligne.hWnd, WM_PASTE, 0, 0
    txttemp.Text = "Vous n'avez pas le droit de copier"
    txttemp.SelStart = 0
    txttemp.SelLength = Len(txttemp)
    Clipboard.Clear
    lignet = Left(Right(Ligne.TextRTF, Len(Ligne.TextRTF) - 115), Len(Ligne.TextRTF) - 125)
Form_Resize
Me.Show
OpenText (txpfile)
End Sub
Function PAKExtract(PAKFile As String, FileToExtract As String, DestinationFile As String) As Boolean
Dim BytesExtract As String
Dim Offset As Long
Dim Size As Long
Dim Name As String

    If PAKValid(PAKFile) = True Then
        If FileExist(DestinationFile) = True Then Kill DestinationFile
        FileNumber = FreeFile
        Open PAKFile For Binary As FileNumber
            Get FileNumber, 7, FileListStart
        
            If FileListStart = 0 Then
                PAKExtract = False
                Close FileNumber
                Exit Function
            Else

                Do
                    Get FileNumber, FileListStart, Offset
                    FileListStart = FileListStart + 4
                
                    Get FileNumber, FileListStart, Size
                    FileListStart = FileListStart + 4
                
                    Name = String$(255, Chr$(0))
                    Get FileNumber, FileListStart, Name
                    Name = Mid(Name, 1, InStr(1, Name, Chr$(0)) - 1)
                    FileListStart = FileListStart + Len(Name) + 1
                
                    If Name = "" Or Offset = 0 Or Size = 0 Then
                        PAKExtract = False
                        Close FileNumber
                        Exit Function
                    ElseIf LCase(Name) = LCase(FileToExtract) Then
                        DestinationNumber = FreeFile
                        Open DestinationFile For Binary As DestinationNumber
                            If Size > 100000 Then
                                
                                Position = -1000000
                                
                                Do
                                    
                                    Position = Position + 1000000
                                    If Position + 999999 > Size Then
                                        BytesExtract = String(Size - Position, Chr$(0))
                                        
                                    Else
                                        BytesExtract = String(1000000, Chr$(0))
                                       
                                    End If
                                    Get FileNumber, Position + Offset, BytesExtract
                                    Put DestinationNumber, Position + 1, BytesExtract
                                Loop Until Position + 999999 >= Size
                            Else
                                BytesExtract = String(Size, Chr$(0))
                                Get FileNumber, Offset, BytesExtract
                                Put DestinationNumber, 1, BytesExtract
                            End If
                        Close DestinationNumber
                        Close FileNumber
                        PAKExtract = True
                        Exit Function
                    End If
                Loop Until FileListStart > LOF(FileNumber)
            End If
        Close FileNumber
        PAKExtract = False
    Else
        PAKExtract = False
        Exit Function
    End If
End Function
Function PAKValid(PAKFileName As String) As Boolean
Dim Header As String
Header = String$(6, Chr$(0))

If FileExist(PAKFileName) = False Then
    PAKValid = False
    Exit Function
Else
    FileNumber = FreeFile
    Open PAKFileName For Binary As FileNumber
        Get FileNumber, 1, Header
        If Header = "TEXTXP" Then
            PAKValid = True
        Else
            PAKValid = False
        End If
    Close FileNumber
End If
End Function
Public Sub OpenText(filename As String)
Dim ptext As String
'On Error Resume Next
If FileExist(filename) = False Then End
PAKExtract filename, "Info", App.Path & "\Info.ini"
Dim a$
Open App.Path & "\Info.ini" For Input As #1
Do Until EOF(1)
Input #1, a$
info.AddItem a$
Loop
Close 1
Kill App.Path & "\Info.ini"
'TXPType = Converter(info.List(0))
If info.ListCount <> 16 Then
MsgBox "Ce fichier n'a pas été créé avec la dernière version de Texte XP. Certaines fonctions ne seront pas disponibles."
End If
If info.ListCount = 16 And info.List(15) = "1" Then
setc.Clear
PAKExtract filename, "Couleurs", App.Path & "\col.txt"

Open App.Path & "\col.txt" For Input As #1
Do Until EOF(1)
Input #1, a$
setc.AddItem a$
Loop
Close 1
Kill App.Path & "\col.txt"

Me.BackColor = setc.List(0)
Shape3.BorderColor = setc.List(13)
Shape4.BorderColor = setc.List(14)
Picture1.BackColor = setc.List(14)
Label2.BackColor = setc.List(14)
Label2.ForeColor = setc.List(15)
Label3.ForeColor = setc.List(15)
'tTextColor = setc.List(16)
Shape1.BorderColor = setc.List(16)
Shape2.BorderColor = setc.List(16)
Shape2.BackColor = setc.List(14)
End If

'fProps.titre.Text = Converter(info.List(1))
'fProps.Auteur.Text = Converter(info.List(2))
'fProps.Theme.Text = Converter(info.List(3))
'fProps.Société.Text = Converter(info.List(4))
'fProps.Comment.Text = Converter(info.List(5))
Rt.BackColor = info.List(6)
'fProps.Col4.BackColor = info.List(7)
'fProps.Col6.BackColor = info.List(8)
'fProps.Col.BackColor = Rt.BackColor
    
    Me.Enabled = False
If FileExist(App.Path & "\Texte.tm1") = True Then Kill App.Path & "\Texte.tm1"
PAKExtract filename, "TexteXP", App.Path & "\Texte.tm1"

Dim tsize As String
If FileExist(App.Path & "\Adler32.txt") = True Then Kill App.Path & "\Adler32.txt"
PAKExtract filename, "Adler32", App.Path & "\Adler32.txt"
If FileExist(App.Path & "\Adler32.txt") = False Then
MsgBox "Ce Texte ne dispose pas de certificat de validité. Il se peut que son contenu soit incorrect."
ElseIf FileExist(App.Path & "\Adler32.txt") = True Then
Open App.Path & "\Adler32.txt" For Input As #2
Input #2, a$
tsize = a$
Close 2
Kill App.Path & "\Adler32.txt"
If tsize <> AdlerFromFile(App.Path & "\Texte.tm1") Then
    MsgBox "Le certificat du fichier semble être incorrect"
    Me.Enabled = True
    End
End If
End If
If Converter(info.List(1)) <> "" Or Converter(info.List(2)) <> "" Then MsgBox "Titre : " & Converter(info.List(1)) & vbCrLf & "Auteur : " & Converter(info.List(2))

If info.List(13) = 0 Then
Call Huffman.DecodeFile(App.Path & "\Texte.tm1", App.Path & "\Texte.tm2")
txttemp.LoadFile App.Path & "\Texte.tm2", 1
ptext = txttemp.Text

ElseIf info.List(13) = 1 Then
pass.Text = InputBox("Ce texte est crypté. Veuillez entrer un mot de passe", "Mot de passe")

If pass.Text = "" Then Exit Sub
Key = pass.Text
    If Crypt(info.List(14)) <> "Code" Then
    MsgBox "Le mot de passe est incorrect"
    Me.Enabled = True
    End
    End If
Call Huffman.DecodeFile(App.Path & "\Texte.tm1", App.Path & "\Texte.tm2")
txttemp.LoadFile App.Path & "\Texte.tm2", 1
Key = pass.Text
txttemp.Text = Crypt(txttemp.Text)
ptext = txttemp.Text

ElseIf info.List(13) = 2 Then
Key = Utilisateur
    If Crypt(info.List(14)) <> "Code" Then
    MsgBox "Voux n'êtes pas autorisé à lire ce fichier"
    Enabled = True
    End
    Exit Sub
    End If
Call Huffman.DecodeFile(App.Path & "\Texte.tm1", App.Path & "\Texte.tm2")
txttemp.LoadFile App.Path & "\Texte.tm2", 1
Key = Utilisateur
txttemp.Text = Crypt(txttemp.Text)
ptext = txttemp.Text

End If

Kill App.Path & "\Texte.tm1"
Kill App.Path & "\Texte.tm2"

ptext = Replace(ptext, "<ligne>", lignet)
ptext = Replace(ptext, "\/", vbCrLf & "\par" & vbspace)

Son.Stop
If info.List(10) = "1" Then
PAKExtract filename, "Son", App.Path & "\texte.tmp"
Son.filename = App.Path & "\texte.tmp"
'son.Play
End If
If info.List(11) <> "0" Then
Piccount.Caption = 0
Dim j

For j = 0 To info.List(11) - 1
PAKExtract filename, "Pic" & j, App.Path & "\Pic" & j
If FileExist(App.Path & "\pic" & j) = False Then Piccount.Caption = Piccount.Caption - 1
    pictemp.Picture = LoadPicture(App.Path & "\pic" & j)
    Insertimg2 pictemp.Picture, False
    ptext = Replace(ptext, "<image:" & j & ">", LCase(imgText(j)))
    Kill App.Path & "\pic" & j
Shape2.Width = Shape1.Width * (j + 1) / info.List(11)
DoEvents
Next j

End If
Rt.TextRTF = ptext
Rt.SelStart = 0
Rt.SelLength = Len(Rt.Text)
Rt.SelIndent = 0
Rt.SelLength = 0
Me.Enabled = True
Shape2.Width = 0
    FormatMailAddress Rt
    convertHyperlink Rt, "www."
    convertHyperlink Rt, "http:"
    convertHyperlink Rt, "mailto:"
    Rt.MousePointer = 0
End Sub

Public Property Let Key(ByVal Key As String)
On Error Resume Next
    Dim Longueur As Long, t As Long
    Longueur = Len(Key)
    For i = 0 To 255
        s(i) = i
    Next i
    
    j = 0
    For i = 0 To 255
        j = (j + s(i) + Asc(Mid$(Key, i Mod Longueur + 1, 1))) And 255&
        t = s(i)
        s(i) = s(j)
        s(j) = t
    Next i
    i = 0
    j = 0
End Property

Public Function Crypt(ByVal Param As String) As String
    Dim Longueur As Long, C As Long, t As Long
    Longueur = Len(Param)
    For C = 1 To Longueur
        i = (i + 1) And 255&
        j = (j + s(i)) And 255&
        t = s(i)
        s(i) = s(j)
        s(j) = t
        
        t = (s(i) + s(j)) And 255&
        
        Mid$(Param, C, 1) = Chr$(Asc(Mid$(Param, C, 1)) Xor s(t))
    Next C
    Crypt = Param
End Function

Function FileExist(filename As String) As Boolean
On Error GoTo Erro

If FileLen(filename) <> 0 Then
    FileExist = True
Else
    FileExist = False
End If
Exit Function

Erro:
If Err = 76 Or Err = 53 Then FileExist = False
End Function

Private Sub Form_Resize()
On Error Resume Next
Picture1.Left = Width - Picture1.Width
Picture1.Top = Height - Picture1.Height
If imouse = 0 Then
Rt.Height = Me.Height - Rt.Top - 100
Rt.Width = Me.Width - (Rt.Left * 2)
Shape1.Width = Rt.Width
Shape3.Width = Me.Width
Shape4.Width = Me.Width - 40
Shape3.Height = Me.Height - Shape3.Top
Shape4.Height = Me.Height - Shape4.Top - 10
Label2.Width = Me.Width
Label3.Left = Me.Width - Label3.Width - 10
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Son.Stop
If FileExist(App.Path & "\texte.tmp") = True Then Kill App.Path & "\texte.tmp"
Unload Me
End Sub
Function Insertimg2(aStdPic As StdPicture, copie As Boolean)
  On Error Resume Next
    Dim hMetaDC     As Long
    Dim hMeta       As Long
    Dim hPicDC      As Long
    Dim hOldBmp     As Long
    Dim aBMP        As BITMAP
    Dim aSize       As Size
    Dim aPt         As POINTAPI
    Dim filename    As String
    Dim screenDC    As Long
    Dim headerStr   As String
    Dim retStr      As String
    Dim byteStr     As String
    Dim Bytes()     As Byte
    Dim filenum     As Integer
    Dim numBytes    As Long
    Dim i           As Long
    
    filename = getTempName("WMF")
    hMetaDC = CreateMetaFile(filename)
    SetMapMode hMetaDC, MM_ANISOTROPIC
    SetWindowOrgEx hMetaDC, 0, 0, aPt
    GetObject aStdPic.Handle, Len(aBMP), aBMP
    SetWindowExtEx hMetaDC, aBMP.bmWidth, aBMP.bmHeight, aSize
    SaveDC hMetaDC
    screenDC = GetDC(0)
    hPicDC = CreateCompatibleDC(screenDC)
    ReleaseDC 0, screenDC
    hOldBmp = SelectObject(hPicDC, aStdPic.Handle)
    BitBlt hMetaDC, 0, 0, aBMP.bmWidth, aBMP.bmHeight, hPicDC, 0, 0, vbSrcCopy
    SelectObject hPicDC, hOldBmp
    DeleteDC hPicDC
    DeleteObject hOldBmp
    RestoreDC hMetaDC, True
    hMeta = CloseMetaFile(hMetaDC)
    DeleteMetaFile hMeta
    
    headerStr = "{\rtf1\ansi"
    headerStr = headerStr & _
                "{\pict\picscalex100\picscaley100" & _
                "\picw" & aStdPic.Width & "\pich" & aStdPic.Height & _
                "\picwgoal" & aBMP.bmWidth * Screen.TwipsPerPixelX & _
                "\pichgoal" & aBMP.bmHeight * Screen.TwipsPerPixelY & _
                "\wmetafile8"
    numBytes = FileLen(filename)
    ReDim Bytes(1 To numBytes)
    filenum = FreeFile()
    Open filename For Binary Access Read As #filenum
    Get #filenum, , Bytes
    Close #filenum
    byteStr = String(numBytes * 2, "0")
    For i = LBound(Bytes) To UBound(Bytes)
        If Bytes(i) > &HF Then
            Mid$(byteStr, 1 + (i - 1) * 2, 2) = Hex$(Bytes(i))
        Else
            Mid$(byteStr, 2 + (i - 1) * 2, 1) = Hex$(Bytes(i))
        End If
    Next i
    retStr = headerStr & " " & byteStr & "}"
    retStr = retStr & "}"
txttemp.Text = ""
txttemp.TextRTF = retStr
txttemp.Text = txttemp.TextRTF
txttemp.Text = Right(txttemp.Text, Len(txttemp.Text) - InStr(txttemp.Text, "{\pict\") + 1)
txttemp.Text = Right(txttemp.Text, Len(txttemp.Text) - InStr(1, txttemp.Text, vbCrLf) - 1)
txttemp.Text = Left(txttemp.Text, InStr(txttemp.Text, "}") - 2)
      
      If TXPreader.Piccount.Caption < 0 Then TXPreader.Piccount.Caption = 0
      imgText(TXPreader.Piccount.Caption) = txttemp.Text
TXPreader.Piccount.Caption = TXPreader.Piccount.Caption + 1

If copie = True Then TXPreader.Rt.SelRTF = retStr
    'StdPicAsRTF = retStr
    On Local Error Resume Next
    If Dir(filename) <> "" Then Kill filename
End Function

Function Utilisateur() As String
    Dim Ch As String
    Dim a As Long
    Dim b As Long

    a = 199
    Ch = String$(200, 0)
    b = GetUserName(Ch, a)
    If b <> 0 Then Utilisateur = Left$(Ch, a) Else Utilisateur = ""
End Function

Private Sub Label2_DblClick()
If Me.WindowState = 0 Then
Me.WindowState = 2
Else
Me.WindowState = 0
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Label3_Click()
Dim re As Integer
re = MsgBox("Etes-vous sur de vouloir quitter?", vbYesNo)
If re = 6 Then Unload Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imouse = 1
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If imouse = 1 Then
Me.Width = x + Width
Me.Height = y + Height
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imouse = 0
Form_Resize
End Sub

Public Function FormatMailAddress(box) As String
    Dim Pos As Long
    Dim chARond As String
    Dim Txt As String
    Dim txtlen As Long
    Dim pos_start As Long
    Dim pos_mijloc As Long
    Dim pos_end As Long
    Dim dom As Boolean
    
    Dim MailAddressStart As Long
    
    Dim mailaddress As String
    Txt = box.Text
    
    Pos = box.Find("@", 1, Len(Txt))
    
    If Pos <= 0 Then
         Exit Function
    End If
    
While Pos > 0
    dom = False
    For pos_start = Pos To 1 Step -1
        chARond = Mid$(Txt, pos_start, 1)
        If chARond = Chr(32) Or chARond = vbCr Or chARond = vbLf Or chARond = vbNewLine Then Exit For
    Next pos_start
    pos_start = pos_start + 1

    txtlen = Len(Txt)
    For pos_end = Pos To txtlen
        chARond = Mid$(Txt, pos_end, 1)
            If dom = True Then
                If (Asc(chARond) < 65 Or Asc(chARond) > 91) And (Asc(chARond) < 97 Or Asc(chARond) > 122) Then Exit For
            End If
            If chARond = "." Then dom = True
    Next pos_end
    
    If pos_start <= pos_end Then
        mailaddress = Mid(box.Text, pos_start, pos_end - pos_start)
        mailaddress = Replace(mailaddress, ">", "")
        mailaddress = Replace(mailaddress, "<", "")
        MailAddressStart = box.Find(mailaddress, pos_start - 1)
        
        box.SelColor = info.List(7)
        box.SelUnderline = True
        box.SelStart = MailAddressStart + 2
    End If
    Pos = box.Find("@", Pos + 1, Len(Txt))
Wend
End Function
Public Sub convertHyperlink(box, keyWord As String)
    Dim hypStart As Long
    Dim befor As String
    Dim after As String
    Dim cuvantAddress As String
    Dim hypEnd As Long
      
    Dim separator1 As String
    Dim separator2 As String
    
    hypStart = box.Find(keyWord, 1, Len(box.Text))
    While hypStart > 0
            separator1 = InStr(hypStart + 1, box.Text, vbCr)
            separator2 = InStr(hypStart + 1, box.Text, Chr(32))
            hypEnd = separator1
            If separator1 > separator2 Then hypEnd = separator2
            If separator2 = 0 Then hypEnd = separator1
            If separator1 = 0 Then hypEnd = Len(box.Text) + 1
        cuvantAddress = Mid(box.Text, hypStart + 1, hypEnd - hypStart - 1)
        
        box.SelStart = hypStart
        box.Find cuvantAddress, hypStart
        box.SelUnderline = True
        box.SelColor = info.List(7)
        box.SelStart = hypStart + 1
        
        hypStart = box.Find(keyWord, hypStart + 1, Len(box.Text))
        
    Wend
End Sub

Public Function GetHyperlink(rch, x As Single, y As Single) As String
   ' On Error Resume Next
    Dim pt As POINTAPI
    Dim Pos As Long
    Dim Ch As String
    Dim Txt As String
    Dim txtlen As Long
    Dim pos_start As Long
    Dim pos_mijloc As Long
    Dim pos_end As Long
    
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY

    Pos = SendMessage(rch.hWnd, &HD7, 0&, pt)
    If Pos <= 0 Then
        Exit Function
    End If
    Txt = rch.Text
    For pos_start = Pos To 1 Step -1
        If Mid$(Txt, pos_start + 1, 1) = Chr(13) Then
            Me.Rt.ToolTipText = ""
            Rt.MousePointer = 0
            Exit Function
        End If
        Ch = Mid$(Txt, pos_start, 1)
        If Ch = Chr(32) Or Ch = vbCr Or Ch = vbLf Or Ch = vbNewLine Then Exit For
    Next pos_start
    pos_start = pos_start + 1

    txtlen = Len(Txt)
    For pos_end = Pos To txtlen
        Ch = Mid$(Txt, pos_end, 1)
    If Ch = Chr(32) Or Ch = vbCr Then Exit For
    Next pos_end
    pos_end = pos_end - 1

    If pos_start <= pos_end Then _
        GetHyperlink = Mid$(Txt, pos_start, pos_end - pos_start + 1)
        
        If Left(GetHyperlink, 5) = "http:" Or Left(GetHyperlink, 4) = "www." Or Left(GetHyperlink, 7) = "mailto:" Then
            Rt.MouseIcon = Cur.Picture
            Rt.MousePointer = vbCustom
            If Left(GetHyperlink, 7) <> "mailto:" Then
            Else
            End If
        ElseIf InStr(1, GetHyperlink, "@") > 0 Then
            Rt.MouseIcon = Cur.Picture
            Rt.MousePointer = vbCustom
        Else
            Rt.MousePointer = 0
        End If
End Function

Private Sub Rt_KeyUp(KeyCode As Integer, Shift As Integer)
    SendMessage txttemp.hWnd, WM_COPY, vbNull, vbNull
End Sub

Private Sub Rt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    htxt = GetHyperlink(Rt, x, y)
End Sub
Private Sub Rt_Click()
    Dim lngRet
    If Left(htxt, 5) = "http:" Or Left(htxt, 7) = "mailto:" Or Left(htxt, 4) = "www." Then lngRet = ShellExecute(0&, "Open", htxt, "", vbNullString, 1)
    If InStr(1, htxt, "@") > 0 Then lngRet = ShellExecute(0&, "Open", "mailto:" + htxt, "", vbNullString, 1)
End Sub

Public Function Adler32(ByVal lngAdler32 As Long, ByRef bArrayIn() As Byte, _
    ByVal dblLength As Double) As Long
    
    Dim intPos As Integer
    Dim lngPosInArray As Long
    Dim lngLengthRemaining As Long
    Dim dblLow As Double
    Dim dblHigh As Double
    
    If lngAdler32 <> 0 Then
        dblLow = lngAdler32 And 65535
        dblHigh = RShiftNoRound(lngAdler32, 16) And 65535
    End If
    
    If UBound(bArrayIn) < LBound(bArrayIn) Then
        Adler32 = 1
    Else
        lngLengthRemaining = dblLength
        
        Do While (lngLengthRemaining > 0)
            If lngLengthRemaining < NMAX Then
                intPos = lngLengthRemaining
                lngLengthRemaining = 0
            Else
                intPos = NMAX
                lngLengthRemaining = lngLengthRemaining - (NMAX + 1)
            End If
            
            Do
                dblLow = dblLow + bArrayIn(lngPosInArray)
                dblHigh = dblHigh + dblLow
                
                lngPosInArray = lngPosInArray + 1
                intPos = intPos - 1
            Loop While intPos >= 0
            dblLow = Modulus(dblLow, BASE)
            dblHigh = Modulus(dblHigh, BASE)
            
        Loop
        Adler32 = LShift4Byte(dblHigh, 16) Or dblLow
    End If
End Function
Private Function Modulus(ByVal dblValue As Double, ByVal dblModValue As Double) As Double
    Modulus = dblValue - (dblModValue * Fix(dblValue / dblModValue))
End Function
Private Function LShift4Byte(ByVal pnValue As Double, ByVal pnShift As Integer) As Long
    Dim lngMask As Long
    If pnValue And (2 ^ (31 - pnShift)) Then
        lngMask = &H80000000
    End If
    LShift4Byte = ((pnValue And ((2 ^ (31 - pnShift)) - 1)) * (2 ^ pnShift)) Or lngMask
End Function

Private Function RShiftNoRound(ByVal pnValue As Double, ByVal pnShift As Integer) As Double
    RShiftNoRound = Int(pnValue / (2 ^ pnShift))
End Function
Function AdlerFromFile(ByVal strFilePath As String) As String
    Dim bArrayFile() As Byte
    Dim lngAdler32 As Long
    
    Dim lngChunkSize As Long
    Dim lngSize As Long
    
    lngSize = FileLen(strFilePath)
    lngChunkSize = CHUNK_SIZE
    
    If lngSize <> 0 Then
        
        Open strFilePath For Binary Access Read As #1
        
        Do While Seek(1) < lngSize
            
            If (lngSize - Seek(1)) > lngChunkSize Then
                Do While Seek(1) < (lngSize - lngChunkSize)
                    ReDim bArrayFile(lngChunkSize - 1)
                    Get #1, , bArrayFile()
                    lngAdler32 = Adler32(lngAdler32, bArrayFile, UBound(bArrayFile))
                Loop
            Else
                ReDim bArrayFile(lngSize - Seek(1))
                Get #1, , bArrayFile()
                lngAdler32 = Adler32(lngAdler32, bArrayFile, UBound(bArrayFile))
            End If
            
        Loop
        
        Close #1
        
        AdlerFromFile = Right$("00000000" & Hex$(lngAdler32), 8)
    Else
        AdlerFromFile = "00000001"
    End If
End Function

