VERSION 5.00
Begin VB.Form Dem 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Dem.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1220
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   1560
      TabIndex        =   3
      Top             =   820
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   2970
      TabIndex        =   8
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   130
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   200
      Left            =   3000
      TabIndex        =   6
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Validate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   240
      TabIndex        =   5
      Top             =   2000
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   2970
      Picture         =   "Dem.frx":21C42
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   150
      Picture         =   "Dem.frx":23124
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Login :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Winux  v:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Dem"
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

Dim cod As String, vers As String

Private Sub Form_Load()
'vérification de la présence de la license
On Error GoTo gpl
Open (App.Path & "\root\system\entete.txt") For Input As #1
Close #1
Open (App.Path & "\root\system\GNU General Public License.txt") For Input As #1
Close #1
'récupération de la version
Open (App.Path & "\root\system\version.ini") For Input As #1
Input #1, vers
Close #1
Label1.Caption = Label1.Caption & vers
Exit Sub
'si pas de licence alors on ferme
gpl:    Unload Me
        End
End Sub

Private Sub Label6_Click()
On Error GoTo 1
If Text1.Text = "root" Then
    Open (App.Path & "\" & Text1.Text & "\" & Text1.Text & ".txt") For Input As #1
Else
    Open (App.Path & "\home\" & Text1.Text & "\" & Text1.Text & ".txt") For Input As #1
End If
Input #1, cod
Close #1
cod = Cryptage(cod, 2002, 1)
If cod = Text2.Text Then
    login = Text1.Text
    Bureau.Show
    Unload Me
Else
    MsgBox "Password incorrect !"
    Text1.Text = ""
    Text2.Text = ""
End If
Exit Sub
1: MsgBox "Username incorrect !"
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Visible = True
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Visible = False
End Sub

Private Sub Label7_Click()
Unload Me
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Visible = True
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Visible = False
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Shift And KeyCode = 13 Then
    Image1.Visible = True
    Label6_Click
    KeyCode = 0
End If
End Sub
