VERSION 5.00
Begin VB.Form Explorateur 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   ClientHeight    =   4455
   ClientLeft      =   2145
   ClientTop       =   1980
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Explorateur.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin Winux.xpcmdbutton xpcmdbutton4 
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      ToolTipText     =   "Delete"
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Winux.xpcmdbutton xpcmdbutton3 
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      ToolTipText     =   "Past"
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Past"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Winux.xpcmdbutton xpcmdbutton2 
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      ToolTipText     =   "Copy"
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Copy"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Winux.xpcmdbutton xpcmdbutton1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Cut"
      Top             =   360
      Width           =   855
      _ExtentX        =   2990
      _ExtentY        =   450
      Caption         =   "Cut"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   2920
      TabIndex        =   1
      Top             =   1200
      Width           =   4650
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2790
      Left            =   105
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Titre 
      BackStyle       =   0  'Transparent
      Caption         =   "Office"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   230
      Left            =   360
      TabIndex        =   5
      Top             =   55
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   1320
      TabIndex        =   3
      Top             =   830
      Width           =   5055
   End
   Begin VB.Label close 
      BackStyle       =   0  'Transparent
      Height          =   260
      Left            =   7320
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   7360
      Picture         =   "Explorateur.frx":25642
      Top             =   40
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "Explorateur"
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

Dim Status, X_Initial, Y_Initial, Dist_Am  ' Pour le déplacement de la fenêtre.
Dim Fichier As String, cheminFichier As String, chemin As String, extension As String, opération As Boolean 'Pour les opérations
'Pour la compression
Dim PIn As String, POut As String, LevelC As Integer, Max As Long, Verif As String

Private Sub Dir1_Change()
If login <> "root" And Left(Dir1.Path, 1) = Left(App.Path, 1) And Len(Dir1.Path) <= (Len(App.Path) + 5) Then
    Dir1.Path = App.Path & "\home\" & login
Else
    File1.Path = Dir1.Path
    File1_Click
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Fichier = File1.filename
cheminFichier = File1.Path
chemin = cheminFichier & "\" & Fichier
On Error GoTo 1
chemin = Mid(chemin, Len(App.Path) + 1, (Len(chemin) - Len(App.Path)))
Label1.Caption = chemin
1: End Sub

Private Sub File1_DblClick()
extension = Right(File1.filename, 3)
If (extension = "exe") Then
    Shell (File1.Path & "\" & File1.filename), vbNormalFocus
ElseIf (extension = "txp") Then
    TXPreader.txpfile = File1.Path & "\" & File1.filename
    TXPreader.Show 0, Bureau
End If
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = 2) And (Label1.Caption <> "") Then
    If (Right(File1.filename, 3) = "wbu") Then
        Menus.archiv.Visible = False
        Menus.resto.Visible = True
        PopupMenu Menus.explo
    Else
        Menus.resto.Visible = False
        Menus.archiv.Visible = True
        PopupMenu Menus.explo
    End If
End If
End Sub

Private Sub File1_PathChange()
Dir1.Path = File1.Path
End Sub

Private Sub Form_Load()
If login = "root" Then
    cheminFichier = App.Path & "\root"
    Explorateur.Dir1.Path = cheminFichier
    Label1.Caption = "\root"
Else
    cheminFichier = App.Path & "\home\" & login
    Explorateur.Dir1.Path = cheminFichier
    Label1.Caption = "\home\" & login
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Status = 1
 X_Initial = x
 Y_Initial = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Status = 1 Then
  Me.Left = Me.Left + x - X_Initial
  Me.Top = Me.Top + y - Y_Initial
 End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 Status = 0
 Dist_Am = 100
 If Me.Left < Dist_Am Then Me.Left = 0
 If Me.Top < Dist_Am Then Me.Top = 0
 If Me.Left + Me.Width > Screen.Width - Dist_Am Then Me.Left = Screen.Width - Me.Width
 If Me.Top + Me.Height > Screen.Height - Dist_Am Then Me.Top = Screen.Height - Me.Height
 If Me.Top < 1000 Then
  Me.Top = 1000
 End If
 If Me.Top > 3600 Then
  Me.Top = 3600
 End If
 If Me.Left < 0 Then
  Me.Left = 0
 End If
 If Me.Left > 7550 Then
  Me.Left = 7550
 End If
End Sub

Private Sub close_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Visible = True
End Sub

Private Sub close_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Visible = False
End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub xpcmdbutton1_Click()
opération = True
End Sub

Private Sub xpcmdbutton2_Click()
opération = False
End Sub

Private Sub xpcmdbutton3_Click()
If opération = False Then
 FileCopy (cheminFichier & "\" & Fichier), (Explorateur.File1.Path & "\" & Fichier)
 Explorateur.File1.Refresh
ElseIf opération = True Then
 FileCopy (cheminFichier & "\" & Fichier), (Explorateur.File1.Path & "\" & Fichier)
 Kill (cheminFichier & "\" & Fichier)
 Explorateur.File1.Refresh
End If
End Sub

Private Sub xpcmdbutton4_Click()
Kill (cheminFichier & "\" & Fichier)
End Sub
