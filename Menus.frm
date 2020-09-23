VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Menus 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog ComD 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu explo 
      Caption         =   "Search"
      Begin VB.Menu coup 
         Caption         =   "Cut"
      End
      Begin VB.Menu cop 
         Caption         =   "Copy"
      End
      Begin VB.Menu col 
         Caption         =   "Past"
      End
      Begin VB.Menu esp1 
         Caption         =   "-"
      End
      Begin VB.Menu comp 
         Caption         =   "Compress"
      End
      Begin VB.Menu decomp 
         Caption         =   "Decompresser"
      End
      Begin VB.Menu esp2 
         Caption         =   "-"
      End
      Begin VB.Menu archiv 
         Caption         =   "Archive"
      End
      Begin VB.Menu resto 
         Caption         =   "Restore"
         Visible         =   0   'False
      End
      Begin VB.Menu searchmnu 
         Caption         =   "Look for..."
      End
   End
   Begin VB.Menu buro 
      Caption         =   "Office"
      Begin VB.Menu fon 
         Caption         =   "Change the bottom"
      End
   End
End
Attribute VB_Name = "Menus"
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

'Pour la compression
Dim PIn As String, POut As String, LevelC As Integer, Max As Long, Verif As String

Private Sub archiv_Click()
If DCencour = True Then
    MsgBox "Il y a déjà une opération en cours ! Patientez."
    Exit Sub
End If
DCencour = True
LevelC = 9
PIn = Explorateur.File1.Path & "\" & Explorateur.File1.filename
POut = App.Path & "\home\" & login & "\wbu\" & Explorateur.File1.filename & ".wbu"
Open (App.Path & "\home\" & login & "\wbu\" & Explorateur.File1.filename & ".wbup") For Output As #1
    Print #1, PIn
Close #1
Result = Compression.CompressFile(PIn, POut, LevelC)
Explorateur.File1.Refresh
End Sub

Private Sub col_Click()
If opération = False Then
 FileCopy (cheminFichier & "\" & Fichier), (Explorateur.File1.Path & "\" & Fichier)
 Explorateur.File1.Refresh
ElseIf opération = True Then
 FileCopy (cheminFichier & "\" & Fichier), (Explorateur.File1.Path & "\" & Fichier)
 Kill (cheminFichier & "\" & Fichier)
 Explorateur.File1.Refresh
End If
End Sub

Private Sub comp_Click()
If DCencour = True Then
    MsgBox "Il y a déjà une opération en cours ! Patientez."
    Exit Sub
End If
DCencour = True
LevelC = 9
PIn = Explorateur.File1.Path & "\" & Explorateur.File1.filename
POut = PIn & "_"
Result = Compression.CompressFile(PIn, POut, LevelC)
Explorateur.File1.Refresh
End Sub

Private Sub cop_Click()
opération = False
End Sub

Private Sub coup_Click()
opération = True
End Sub

Private Sub decomp_Click()
If DCencour = True Then
    MsgBox "Il y a déjà une opération en cours ! Patientez."
    Exit Sub
End If
DCencour = True
PIn = Explorateur.File1.Path & "\" & Explorateur.File1.filename
Verif = Right(PIn, 1)
If Verif <> "_" Then
MsgBox ("Ce fichier n'est pas un fichier compressé au format Softzip !")
GoTo 1
End If
Max = Len(PIn) - 1
POut = Mid(PIn, 1, Max)
Result = Compression.DecompressFile(PIn, POut)
1: End Sub

Private Sub fon_Click()
ComD.filename = ""
ComD.InitDir = App.Path & "\" & login & "\"
ComD.Filter = "Fichier image (*.jpg)|*.jpg|Fichier image (*.jpeg)|*.jpeg|Fichier image (*.bmp)|*.bmp|"
ComD.ShowOpen
If ComD.filename <> "" Then
    Bureau.Image1.Picture = LoadPicture(ComD.filename)
    Call UserBase(False, 0, ComD.filename, True)
Else
    Call UserBase(False, 0, "a", True)
End If
End Sub

Private Sub resto_Click()
If DCencour = True Then
    MsgBox "Il y a déjà une opération en cours ! Patientez."
    Exit Sub
End If
DCencour = True
PIn = Explorateur.File1.Path & "\" & Explorateur.File1.filename
Open (App.Path & "\home\" & login & "\wbu\" & Explorateur.File1.filename & ".wbup") For Input As #1
    Input #1, POut
Close #1
Max = Len(POut) - 4
POut = Mid(POut, 1, Max)
Result = Compression.DecompressFile(PIn, POut)
End Sub

Private Sub searchmnu_Click()
Cherche.Show 0, Bureau
Cherche.Dir1.Path = Explorateur.File1.Path
End Sub
