VERSION 5.00
Begin VB.Form Arrêt 
   Appearance      =   0  'Flat
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000016&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000016&
      Caption         =   "Change User Sesion"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "Shut Down"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Arrêt"
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

Private Sub Command1_Click()
Unload Menus
Unload Explorateur
Unload Cherche
Unload Home
Unload TXPreader
Unload Users
Unload Infos
Unload Bureau
Unload Me
End
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
