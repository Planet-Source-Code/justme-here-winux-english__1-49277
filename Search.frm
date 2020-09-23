VERSION 5.00
Begin VB.Form Cherche 
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
   Picture         =   "search.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin Winux.xpcmdbutton Command2 
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Enabled         =   0   'False
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
   Begin Winux.xpcmdbutton Command1 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Search"
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
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   900
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   2880
      TabIndex        =   5
      Top             =   1320
      Width           =   4455
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   3240
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label TitleBar 
      BackStyle       =   0  'Transparent
      Height          =   280
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Titre 
      BackStyle       =   0  'Transparent
      Caption         =   "Search for..."
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
      TabIndex        =   2
      Top             =   55
      Width           =   2535
   End
   Begin VB.Label close 
      BackStyle       =   0  'Transparent
      Height          =   260
      Left            =   7320
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   7360
      Picture         =   "search.frx":6F642
      Top             =   40
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "Cherche"
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
'Déclarations pour la recherche de fichiers
Dim Find As New clsFind  ' Déclare une nouvelle instance de la classe
Dim searchpath(32000) As String  'tableau pour les chemins des fichiers

Private Sub Command1_Click() ' Rechercher
If InStr(1, Text1.Text, ".", 1) = 0 Then
Text1.Text = Text1.Text & ".*"
End If
'If Right$(Text1.Text, 2) <> ".*" Or Left$(Right$(Text1.Text, 4), 1) <> "." Then
'    Text1.Text = Text1.Text & ".*"  'Pour améliorer la recherche
'End If
' Variables
Dim i As Long  ' Effectuer la boucle pour récupérer les fichiers trouver
      Command1.Enabled = False  ' Met le bouton(Lancer la recherche) à Disabled
      Command2.Enabled = True  ' Met le bouton(Stopper la recherche) à Enabled
      ' Affiche zéro pour commencé
      Label1.Caption = "Nombres de fichiers trouver: 0"
      List1.Clear  ' Vide la ListBox
      ' Indique à la classe si elle doit rechercher dans les sous-répertoires, ...
      Find.WithSubFolder = True
      Find.Path = Dir1.Path  ' Indique le répertoire de recherche
      Find.FileType = Text1.Text  ' Indique le type de fichier à rechercher
      Find.Search  ' Lance la recherche
      DoEvents  ' On respire un peu :)
      ' Renvoie zéro(0) si aucun fichier n'à été trouver
      If Find.NumFiles > 0 Then
            ' Récupère tous les fichiers trouver un à un
            For i = 1 To Find.NumFiles
                  ' Ajoute le nom des fichiers seulement
                  List1.AddItem Find.GetFileTitle(i)
                  ' Pour ajouter le chemin d'accès complet
                  ' List1.AddItem Find.GetFile(I)
                  '
                  ' Seulement le nom du répertoire ou il ce trouve
                  searchpath(i) = Find.GetFilePath(i)
                  ' List1.AddItem Find.GetFilePath(I)
            Next i
            ' Affiche le nombres de fichiers trouver
            Label1.Caption = "Nombres de fichiers trouver: " & Find.NumFiles
      'Else
            ' Sinon on informe l'utilisateur que l'on à pas trouver de fichiers
            'MsgBox "Aucun fichier n'à été trouver !", vbOKOnly + vbInformation, "Terminer"
      End If
      ' Décharge de la mémoire
      Set Find = Nothing
      List1.Refresh  ' Refresh(met à jour) la ListBox
      Command1.Enabled = True  ' Met le bouton(Lancer la recherche) à Enabled
      Command2.Enabled = False  ' Met le bouton(Stopper la recherche) à Disabled
End Sub

Private Sub Command2_Click()
    Find.Cancel  ' Stop la recherche
    Command1.Enabled = True  ' Met le bouton(Lancer la recherche) à Enabled
    Command2.Enabled = False  ' Met le bouton(Stopper la recherche) à Disabled
End Sub

Private Sub List1_DblClick()
Explorateur.Show 0, Bureau
Explorateur.File1.Path = searchpath(List1.ListIndex)
End Sub

Private Sub Text1_Change()
If Trim$(Text1.Text) = "" Then  ' Si le TextBox est vide
    Command1.Enabled = False  ' Met à Disabled
Else
    Command1.Enabled = True  ' Sinon à Enabled
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Shift And KeyCode = 13 Then
    Command1_Click
End If
End Sub

Private Sub Form_Load()
If login = "root" Then
    Dir1.Path = App.Path & "\root"
Else
    Dir1.Path = App.Path & "\home\" & login
End If
End Sub

Private Sub Dir1_Change()
If login <> "root" And Left(Dir1.Path, 1) = Left(App.Path, 1) And Len(Dir1.Path) <= (Len(App.Path) + 5) Then
    Dir1.Path = App.Path & "\home\" & login
End If
End Sub

Private Sub TitleBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Status = 1
 X_Initial = x
 Y_Initial = y
End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Status = 1 Then
  Me.Left = Me.Left + x - X_Initial
  Me.Top = Me.Top + y - Y_Initial
 End If
End Sub

Private Sub TitleBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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
