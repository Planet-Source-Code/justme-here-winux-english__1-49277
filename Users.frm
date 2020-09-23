VERSION 5.00
Begin VB.Form Users 
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   3345
   ClientLeft      =   4305
   ClientTop       =   2730
   ClientWidth     =   6285
   Icon            =   "Users.frx":0000
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin Winux.xpcmdbutton button2 
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Cancel"
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
   Begin Winux.xpcmdbutton button1 
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Validate"
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
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Retype user pass"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1950
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New user pass:"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of the user:"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   750
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      Height          =   180
      Left            =   5840
      Top             =   30
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00755433&
      X1              =   15
      X2              =   6300
      Y1              =   4665
      Y2              =   4665
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00755433&
      X1              =   6240
      X2              =   6240
      Y1              =   0
      Y2              =   4830
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6045
      TabIndex        =   0
      Top             =   30
      Width           =   195
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00755433&
      X1              =   6210
      X2              =   6210
      Y1              =   60
      Y2              =   210
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00755433&
      X1              =   6060
      X2              =   6225
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   6075
      Picture         =   "Users.frx":08CA
      Top             =   60
      Width           =   135
   End
   Begin VB.Shape Shape4 
      Height          =   195
      Left            =   6045
      Top             =   30
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   6060
      Top             =   45
      Width           =   165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00755433&
      X1              =   -45
      X2              =   6270
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B99063&
      FillColor       =   &H009B6F43&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   6285
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   3100
      Left            =   0
      Top             =   240
      Width           =   6285
   End
   Begin VB.Image Image1 
      Height          =   165
      Left            =   5850
      Picture         =   "Users.frx":0A08
      Top             =   40
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Create a new user:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   75
      TabIndex        =   6
      Top             =   10
      Width           =   5940
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   15
      Picture         =   "Users.frx":0BD6
      Top             =   15
      Width           =   6270
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FDA04F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   5835
      Top             =   15
      Width           =   4935
   End
End
Attribute VB_Name = "Users"
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

Dim Status, X_Initial, Y_Initial, Dist_Am
Dim h99 As Integer
Dim w99 As Integer

Private Sub button1_Click()
If (Text1.Text <> "") And (Text2.Text <> "") And (Text3.Text <> "") Then
    If Text2.Text = Text3.Text Then
        On Error GoTo us
        Open (App.Path & "\home\" & Text1.Text & "\" & Text1.Text & ".txt") For Input As #1
        Close #1
        MsgBox "Cet utilisateur existe déjà ! Choisissez un autre nom."
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Exit Sub
        
us:         MkDir (App.Path & "\home\" & Text1.Text)
            MkDir (App.Path & "\home\" & Text1.Text & "\Documents")
            MkDir (App.Path & "\home\" & Text1.Text & "\Poubelle")
            MkDir (App.Path & "\home\" & Text1.Text & "\Wbu")
            Open (App.Path & "\home\" & Text1.Text & "\user.ini") For Output As #1
                Print #1, "a"
            Close #1
            Open (App.Path & "\home\" & Text1.Text & "\" & Text1.Text & ".txt") For Output As #1
                Print #1, Cryptage(Text2.Text, 2002, 0)
            Close #1
            MsgBox ("L'utilisateur " & Text1.Text & " a été créé avec succès.")
            Unload Me
    Else
        MsgBox "Veuillez entrer deux fois le même mot de passe !"
        Text2.Text = ""
        Text3.Text = ""
    End If
Else
    MsgBox "Veuillez remplir tous les champs !"
End If
End Sub

Private Sub button2_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Call remap
End Sub

Private Sub Image1_Click()
If Me.Height <> Image5.Height + 20 Then
    h99 = Me.Height
    w99 = Me.Width
    Me.Height = Image5.Height + 20
    Me.Width = Image5.Width + 10
    Exit Sub
End If
    Me.Height = h99
    Me.Width = w99
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape5.Visible = True
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Status = 1
 X_Initial = x
 Y_Initial = y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Status = 1 Then
  Me.Left = Me.Left + x - X_Initial
  Me.Top = Me.Top + y - Y_Initial
 Else
  Call remap
 End If
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 Status = 0
 Dist_Am = 100
 If Me.Left < Dist_Am Then Me.Left = 0
 If Me.Top < Dist_Am Then Me.Top = 0
 If Me.Left + Me.Width > Screen.Width - Dist_Am Then Me.Left = Screen.Width - Me.Width
 If Me.Top + Me.Height > Screen.Height - Dist_Am Then Me.Top = Screen.Height - Me.Height
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape7.BorderColor
 Line5.BorderColor = Shape7.BorderColor
 Shape7.BorderColor = Value
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape7.BorderColor
 Line5.BorderColor = Shape7.BorderColor
 Shape7.BorderColor = Value
 Unload Me
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Shape4.Visible = False Then
  Call remap
  Shape4.Visible = True
 End If
End Sub

Sub remap()
 If Shape4.Visible = True Then Shape4.Visible = False
 If Shape5.Visible = True Then Shape5.Visible = False
End Sub

