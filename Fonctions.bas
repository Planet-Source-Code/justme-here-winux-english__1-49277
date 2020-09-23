Attribute VB_Name = "ModuleFonct"
'Winux Graphic User Interface for Windows based systems
'Copyright (C) 2002-2003 Winux Team
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'A copy of this licence is available in root\system directory.
'http://www.winux.free.fr or tex_winux@hotmail.com for more details.

'bureau en arrière plan
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'cacher la taskbar
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
'Déclarations pour l'édition de l'user base
Dim userdata(10) As String, isprim As Boolean, j As Long

'Cryptage/décryptage
Public Function Cryptage(ByVal Entree As String, ByVal Cle As Long, ByVal Operation As Long) As String
' Cryptage "Securicrypt" algo by tex
' -3600 < Cle < 3600
' 0 pour crypter / 1 pour décrypter
' appel fonction: résultat = Cryptage(varEntree, varCle, 0 ou 1)
Dim Entree1(20000)
Dim Sortie1(20000)
Dim Algo As Long
Cle = Cle - 1
If Operation = 0 Then 'Boucle de cryptage
For i = 1 To Len(Entree)
Entree1(i) = Mid(Entree, i, 1)
Cle = Cle + 1
Algo = Cos(Cle ^ 5) * 10
Sortie1(i) = Chr(Asc(Entree1(i)) + Algo)
Cryptage = Cryptage + Sortie1(i)
Next i
End If
If Operation = 1 Then 'Boucle de décryptage
For i = 1 To Len(Entree)
Entree1(i) = Mid(Entree, i, 1)
Cle = Cle + 1
Algo = Cos(Cle ^ 5) * 10
Sortie1(i) = Chr(Asc(Entree1(i)) - Algo)
Cryptage = Cryptage + Sortie1(i)
Next i
End If
End Function

'I/O de la base des paramètres de chaque utilisateur
'by tex 07/10/02 17:46:42(GMT)
'status : ok
'utilisation :
'   écriture : Call UserBase(False, 0, ComD.FileName, True)
'   si plusieurs écritures :
'       Call UserBase(False, 0, ComD.FileName, False)
'       Call UserBase(False, 1, ComD.FileName, False)
'       Call UserBase(False, 2, ComD.FileName, True)
'   lecture : pathfond = UserBase(True, 0, 0, 0)
Public Function UserBase(readonly As Boolean, number As Long, corpus As String, isend As Boolean) As Variant
'écriture dans l'user base
If readonly = False Then
    If isprim = True Then
        Open (App.Path & "\" & login & "\user.ini") For Input As #1
        j = 0
        Do
            Input #1, userdata(j)
            j = j + 1
        Loop Until EOF(1)
        Close #1
    End If
    userdata(number) = corpus
    If isend = True Then
        Open (App.Path & "\" & login & "\user.ini") For Output As #1
        j = 0
        Do
            Print #1, userdata(j)
            j = j + 1
        Loop Until EOF(1)
        Close #1
        isprim = True
    End If
Else
'lecture seule dans l'user base
    If login = "root" Then
        Open (App.Path & "\" & login & "\user.ini") For Input As #1
        j = 0
        Do
            Input #1, userdata(j)
            j = j + 1
        Loop Until EOF(1)
        Close #1
        UserBase = userdata(number)
    Else
         Open (App.Path & "\home\" & login & "\user.ini") For Input As #1
        j = 0
        Do
            Input #1, userdata(j)
            j = j + 1
        Loop Until EOF(1)
        Close #1
        UserBase = userdata(number)
    End If
End If
End Function

'initialisation du Bureau avec les paramètres des utilisateurs
'by tex 25/10/02 20:04:42(GMT)
'status : ok
'utilisation :
'   InitUser
Public Sub InitUser()
    Dim pathfond As String 'chemin du fond d'écran
    pathfond = UserBase(True, 0, 0, 0)
    If pathfond = "a" Then
        Bureau.Image1.Picture = Nothing
    Else
        On Error Resume Next
        Bureau.Image1 = LoadPicture(pathfond)
    End If
End Sub
