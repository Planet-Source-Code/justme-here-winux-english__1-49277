VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Home 
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9345
   Icon            =   "Home.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3165
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   5583
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7110
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":0442
            Key             =   "IMAGE1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":1F94
            Key             =   "IMAGE6"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":3AE6
            Key             =   "IMAGE7"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":5638
            Key             =   "IMAGE8"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":718A
            Key             =   "IMAGE2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":8CDC
            Key             =   "IMAGE3"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":A82E
            Key             =   "IMAGE4"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":C380
            Key             =   "IMAGE5"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1365
      Left            =   45
      TabIndex        =   1
      Top             =   3150
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   2408
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4290
      Left            =   6480
      TabIndex        =   2
      Top             =   45
      Width           =   2760
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4515
      Left            =   6435
      Top             =   0
      Width           =   2850
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      FillStyle       =   0  'Solid
      Height          =   4560
      Left            =   6435
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Home"
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

Private Sub RefreshPoste()
Me.Caption = "Post work"
On Error Resume Next
Dim d
Set fs = CreateObject("Scripting.FileSystemObject")
Set dc = fs.Drives
n = 0
ListView1.BackColor = &H80000000
ListView2.BackColor = &H80000000
For Each d In dc
    n = n + 1
    Select Case d.DriveType
        Case 0:
        ListView1.ListItems.Add , "C00" & Str$(n), d.DriveLetter & ":", "IMAGE5"
        Case 1:
        ListView1.ListItems.Add , "C00" & Str$(n), d.DriveLetter & ":", "IMAGE3"
        Case 2:
        If d.VolumeName <> "" Then
            ListView1.ListItems.Add , "C00" & Str$(n), d.DriveLetter & ":" & " [" & d.VolumeName & "]", "IMAGE2"
        Else
             ListView1.ListItems.Add , "C00" & Str$(n), d.DriveLetter & ":", "IMAGE2"
        End If
        Case 3:
        ListView1.ListItems.Add , "C00" & Str$(n), d.DriveLetter & ":", "IMAGE1"
        Case 4:
        ListView1.ListItems.Add , "C00" & Str$(n), d.DriveLetter & ":", "IMAGE1"
        Case 5:
        ListView1.ListItems.Add , "C00" & Str$(n), d.DriveLetter & ":", "IMAGE4"
    End Select
Next
ListView2.ListItems.Add , "CONFIGWINUX", "Configuration WINUX", "IMAGE6"
ListView2.ListItems.Add , "INFOSWINUX", "Information WINUX", "IMAGE5"
End Sub
Private Sub Form_Load()
RefreshPoste
Label1.Caption = ""
End Sub

Private Sub ListView1_Click()
Set fs = CreateObject("Scripting.FileSystemObject")
Set dc = fs.Drives.Item(Left$(ListView1.SelectedItem.Text, 2))
If dc.IsReady = False Then Exit Sub
s = s & "Volume Name : " & dc.VolumeName & vbCrLf & vbCrLf
s = s & "Serial Number : " & dc.SerialNumber & vbCrLf
s = s & "File System : " & dc.FileSystem & vbCrLf
s = s & "Total Size : " & Round(dc.TotalSize / 1024 / 1024, 2) & " Mb." & vbCrLf
s = s & "Free Space : " & Round(dc.FreeSpace / 1024 / 1024, 2) & " Mb. "
s = s & vbCrLf
Label1.Caption = s
End Sub

Private Sub ListView1_dblClick()
On Error GoTo Perif
Explorateur.Dir1.Path = ListView1.SelectedItem.Text
Explorateur.Show 0, Bureau
Exit Sub
Perif: MsgBox "Device not available!  Please insert disk."
End Sub

Private Sub ListView2_DblClick()
If ListView2.SelectedItem.Key = "CONFIGWINUX" Then
    Users.Show 0, Bureau
ElseIf ListView2.SelectedItem.Key = "INFOSWINUX" Then
    Infos.Show 0, Bureau
End If
End Sub
