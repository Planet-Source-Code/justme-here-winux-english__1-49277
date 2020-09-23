VERSION 5.00
Begin VB.Form Infos 
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   4665
   ClientLeft      =   4305
   ClientTop       =   2730
   ClientWidth     =   6285
   Icon            =   "infos.frx":0000
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   480
      Width           =   5775
   End
   Begin Winux.xpcmdbutton button1 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "OK"
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
      Picture         =   "infos.frx":08CA
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
      Height          =   4425
      Left            =   0
      Top             =   240
      Width           =   6285
   End
   Begin VB.Image Image1 
      Height          =   165
      Left            =   5850
      Picture         =   "infos.frx":0A08
      Top             =   40
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Information on Winux:"
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
      TabIndex        =   1
      Top             =   10
      Width           =   5940
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   15
      Picture         =   "infos.frx":0BD6
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
Attribute VB_Name = "Infos"
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
'pour lire le fichier
Dim gpldata(160) As String

Private Sub button1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim j As Integer
'récupération de l'entête
Open (App.Path & "\root\system\entete.txt") For Input As #1
    j = 0
    Do
        Input #1, gpldata(j)
        Text1.Text = Text1.Text & gpldata(j)
        j = j + 1
    Loop Until EOF(1)
Close #1
'récupération de la licence GPL
Open (App.Path & "\root\system\GNU General Public License.txt") For Input As #1
    j = 0
    Do
        Input #1, gpldata(j)
        Text1.Text = Text1.Text & gpldata(j)
        j = j + 1
    Loop Until EOF(1)
Close #1
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

