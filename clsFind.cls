VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Winux Graphic User Interface for Windows based systems
'Copyright (C) 2002-2003 Winux Team
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'A copy of this licence is available in root\system directory.
'http://www.winux.free.fr or tex_winux@hotmail.com for more details.

Option Explicit

' API's
Private Declare Function GetInputState Lib "user32.dll" () As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

' Constantes
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

' Types
Private Type FILETIME
      dwLowDateTime  As Long
      dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
      dwFileAttributes As Long
      ftCreationTime   As FILETIME
      ftLastAccessTime As FILETIME
      ftLastWriteTime  As FILETIME
      nFileSizeHigh    As Long
      nFileSizeLow     As Long
      dwReserved0      As Long
      dwReserved1      As Long
      cFileName        As String * MAX_PATH
      cAlternate       As String * 14
End Type

' Va contenir tous les fichiers trouver
Private Type FindFileStruct
      FileTitle() As String
      FilePath()  As String
      FullPath()  As String
      NumFiles    As Long
End Type

Private StopSearch As Boolean
Private StartPath  As String   ' Répertoire de recherche
Private TypeFile   As String   ' Type de fichier à rechercher
Private SubFolder  As Boolean  ' Rechercher dans les sous-répertoire ou non
Private FileList   As FindFileStruct  ' Liste des fichiers trouver

Sub Search()  ' Lancer la recherche

      StopSearch = False  ' Pour débuter la recherche
      Call Class_Terminate  ' Vide la liste si elle ne les pas
      
      ' Vérifie si un type(extention) à bien été indiquer
      If Trim$(TypeFile) = "" Then TypeFile = "*.*"
      
      ' Si un répertoire de départ à été indiquer
      If Trim$(StartPath) <> "" Then
      
            ' Si on recherche dans les sous-répertoires, ...
            If SubFolder = True Then
            
                  ' Appel la fonction récursive(pour les sous-répertoires, ...)
                  Call SearchDirectory(StartPath)
                  
            Else
            
                  ' Recherche simplement les fichiers dans le dossier indiquer dans StartPath
                  Call FindFile(StartPath)
                  
            End If
            
      End If

End Sub

Private Sub FindFile(ByRef PathToSearch As String)

' Variables
Dim Find     As Long  ' Pour effectuer la recherche
Dim PathName As String  ' Va contenir le nom du dossier dans lequel va s'effectuer la recherche des fichiers
Dim File     As WIN32_FIND_DATA  ' Va contenir les infos des fichiers trouver

      ' Vérifie si un dossier à été indiquer dans la variable
      If Trim$(PathToSearch) = "" Then Exit Sub
      ' Vérifie si un type à été indiquer à la classe
      If Trim$(TypeFile) = "" Then TypeFile = "*.*"
      
      ' Le dossier de recherche
      PathName = AddSlash(PathToSearch) & TypeFile
      
      ' Lance la recherche
      Find = FindFirstFile(PathName, File)
      If Find <> INVALID_HANDLE_VALUE Then  ' S'il à trouver quelque chose
      
            Do  ' Commence la boucle
            
                  ' Si à True on quitte
                  If StopSearch Then Exit Do
                  ' Pour pouvoir stopper la recherche
                  If GetInputState() Then DoEvents
            
                  ' Vérifie si ce n'est pas le répertoire courant DOS ...
                  If (Left$(File.cFileName, 1) <> "." And Left$(File.cFileName, 2) <> "..") Then
                  
                        ' Ajoute le nouveau fichier trouver à la liste
                        Call AddFileToList(StripNulls(File.cFileName), PathToSearch)
                        
                  End If

            ' Recherche le prochain fichier
            Loop While FindNextFile(Find, File)
            
            ' Ferme la recherche
            Call FindClose(Find)
            
      End If

End Sub

Private Sub SearchDirectory(ByRef DirectoryName As String)

' Variables
Dim Find     As Long  ' Pour effectuer la recherche
Dim PathName As String  ' Va contenir le nom du répertoire dans lequel va s'effectuer la recherche des répertoires, sous-répertoires, ...
Dim File     As WIN32_FIND_DATA  ' Va contenir les infos du répertoire trouver

      ' Ajoute une barre oblique si besoin
      DirectoryName = AddSlash(DirectoryName)
      ' Répertoire de recherche
      PathName = DirectoryName & "*.*"
      
      ' Recherche les fichiers dans le répertoire indiquer dans DirectoryName
      Call FindFile(DirectoryName)
      
      ' Lance la recherche des sous-répertoires, ....
      Find = FindFirstFile(PathName, File)
      If Find <> INVALID_HANDLE_VALUE Then  ' S'il à trouver quelque chose
      
            Do  ' Commence la boucle
            
                  ' Si à True on quitte
                  If StopSearch = True Then Exit Do
                  ' Pour pouvoir stopper la recherche
                  If GetInputState Then DoEvents
            
                  ' Vérifie si c'est bien un répertoire
                  If (File.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                  
                        ' Vérifie si c'est un répertoire courant DOS...
                        If (Left$(File.cFileName, 1) <> "." And Left$(File.cFileName, 2) <> "..") Then
                        
                              ' Ré-appel la fonction dans laquel on ce trouve
                              ' pour rechercher les sous-répertoires dans le
                              ' sous-répertoire trouver
                              Call SearchDirectory(DirectoryName & StripNulls(File.cFileName))
                              
                        End If
                        
                  End If
                  
            ' Recherche le prochain sous-répertoire
            Loop While FindNextFile(Find, File)
            
            ' Ferme la recherche
            Call FindClose(Find)
            
      End If

End Sub

Sub Cancel()

      ' Stopper la recherche
      StopSearch = True

End Sub

Property Get NumFiles() As Long

      ' Renvoie le nombres de fichiers trouver
      NumFiles = FileList.NumFiles

End Property

Function GetFilePath(ByVal Index As Long) As String

      If ((Index > 0 And Index <= FileList.NumFiles) And FileList.NumFiles > 0) Then

            ' Renvoie le nom du répertoire du fichier pointer par Index
            GetFilePath = FileList.FilePath(Index)
            
      End If

End Function

Function GetFileTitle(ByVal Index As Long) As String

      If ((Index > 0 And Index <= FileList.NumFiles) And FileList.NumFiles > 0) Then

            ' Renvoie seulement le nom du fichier pointer par Index
            GetFileTitle = FileList.FileTitle(Index)
            
      End If

End Function

Function GetFile(ByVal Index As Long) As String

      If ((Index > 0 And Index <= FileList.NumFiles) And FileList.NumFiles > 0) Then
      
            ' Renvoie le chemin d'accès complet du fichier pointer par Index
            GetFile = FileList.FullPath(Index)
            
      End If

End Function

Property Let FileType(ByVal TypeName As String)

      ' Indiquer le type(extention) de fichier à rechercher à la classe
      If Trim$(TypeName) <> "" Then TypeFile = TypeName

End Property

Property Get FileType() As String

      ' Renvoie le type(extention) de fichier à rechercher
      FileType = TypeFile

End Property

Property Let WithSubFolder(ByVal bValue As Boolean)

      ' Indique si l'on doit rechercher dans les sous-répertoires, ...
      SubFolder = bValue

End Property

Property Get WithSubFolder() As Boolean

      ' Renvoie True si on recherche dans les sous-répertoires, ... sinon False
      WithSubFolder = SubFolder

End Property

Property Let Path(ByVal PathName As String)

      ' Indique le nom du répertoire de recherche
      If Trim$(PathName) <> "" Then StartPath = PathName

End Property

Property Get Path() As String

      ' Renvoie le nom du répertoire de recherche
      Path = StartPath

End Property

Private Sub AddFileToList(ByVal filename As String, ByRef FilePath As String)
' Ajouter un fichier à la liste

      ' Si une des 2 variables sont vide(Null) on quitte
      If (Trim$(filename) = "" Or Trim$(FilePath) = "") Then Exit Sub

      With FileList
      
            ' Ajoute un barre oblique(\) si besoin
            FilePath = AddSlash(FilePath)
            ' Incrémente le nombres de fichiers trouver de un
            .NumFiles = .NumFiles + 1
                  
            ' Redimmensionne les tableaux
            ReDim Preserve .FilePath(1 To .NumFiles)
            ReDim Preserve .FileTitle(1 To .NumFiles)
            ReDim Preserve .FullPath(1 To .NumFiles)
      
            ' Sauvegarde le nom du fichier et le nom du répertoire dans lequel il ce trouve
            .FilePath(.NumFiles) = FilePath  ' Nom du répertoire
            .FileTitle(.NumFiles) = filename ' Nom du fichier seulement
            .FullPath(.NumFiles) = FilePath & filename  ' Chemin d'accès complet au fichier

      End With

End Sub

Sub ClearList()

      ' Pour vider la liste(libère la mémoire par la même occasion)
      Call Class_Terminate

End Sub

Function AddSlash(ByVal Chaine As String) As String
' Ajouter une barre oblique(\) à la fin d'une chaine

      If Len(Chaine) > 0 Then  ' Si la chaine n'est pas vide(Null)
      
            If Right$(Chaine, 1) <> "\" Then  ' S'il n'y à pas de barre oblique à la fin
            
                  AddSlash = Chaine & "\"  ' On l'ajoute
                  
            Else
            
                  AddSlash = Chaine  ' Sinon on renvoie la chaine
                  
            End If
            
      Else
      
            AddSlash = "\"  ' Sinon on renvoie une barre oblique(\)
            
      End If

End Function

Private Sub Class_Initialize()
' Initialisation des variables

      StopSearch = False
      FileList.NumFiles = 0
      SubFolder = True
      StartPath = CurDir
      TypeFile = "*.*"
            
End Sub

Private Sub Class_Terminate()

      If FileList.NumFiles > 0 Then

            ' Libère la mémoire des fichiers trouver(vide la liste)
            Erase FileList.FilePath()
            Erase FileList.FileTitle()
            Erase FileList.FullPath()
      
      End If

End Sub

Private Function StripNulls(ByVal filename As String) As String
' Pour supprimer les caractères null à la fin d'une chaine
Dim i As Integer  ' Va contenir la position du premier caractère null
    
      ' Recherche le premier caractère null
      i = InStr(filename, Chr$(0))
    
      ' S'il en à trouver un
      If i > 0 Then
    
            ' Renvoie les caractères qui ce trouve avant les ou les caractères null
            StripNulls = Left$(filename, i - 1)
    
      Else
      
            ' Sinon on renvoie le chaine(Filename)
            StripNulls = filename
            
      End If

End Function

'**************************************************************************'
'*                                                                        *'
'*                          D'autres fonctions                            *'
'*                                                                        *'
'**************************************************************************'

Function RemoveSlash(ByVal Chaine As String) As String
' Enlever une barre oblique(\) à la fin d'une chaine

      If Len(Chaine) > 0 Then  ' Si la chaine n,est pas vide(Null)
      
            If Right$(Chaine, 1) = "\" Then  ' S'il y à une barre oblique à la fin de la chaine
            
                  RemoveSlash = Left$(Chaine, Len(Chaine) - 1)  ' On l'enlève
                  
            Else
            
                  RemoveSlash = Chaine  ' Sinon on renvoie la chaine
                  
            End If
            
      Else
      
            RemoveSlash = ""  ' Sinon on renvoie rien
            
      End If

End Function

Function GetFileName(ByVal FullPath As String) As String
' Récupérer le nom du fichier d'un chemin d'accès complet

Dim i As Long

      If Len(FullPath) > 0 Then  ' Si la chaine n'est pas vide(Null)
      
            ' Passe les caractères un à un à partir de la fin
            For i = Len(FullPath) To 1 Step -1
            
                  ' Si on est arriver sur une barre oblique(\) on quitte la boucle
                  If Mid$(FullPath, i, 1) = "\" Then Exit For
                  
            Next i
            
            ' Renvoie le nom du fichier seulement
            GetFileName = Mid$(FullPath, i + 1)

      Else
           
            ' Sinon on renvoie rien
            GetFileName = ""
            
      End If

End Function

Function GetPath(ByRef FullPath As String) As String
' Récupérer le nom du répertoire d'un chemin d'accès complet à un fichier

Dim i As Long

      If Len(FullPath) > 0 Then  ' Si la chaine n'est pas vide(Null)
      
            ' Enlève la barre oblique(\) s'il y en à une à la fin
            FullPath = RemoveSlash(FullPath)
      
            ' Passe tous les caractères un à un à partir de la fin
            For i = Len(FullPath) To 1 Step -1
            
                  ' Si on est arriver sur une barre oblique(\) on quitte la boucle
                  If Mid$(FullPath, i, 1) = "\" Then Exit For
                  
            Next i
            
            ' Renvoie le nom du répertoire seulement
            GetPath = AddSlash(Mid$(FullPath, 1, i))
            
      Else
      
            ' Sinon on renvoie rien
            GetPath = ""
            
      End If

End Function

Function RemoveExtention(ByVal filename As String) As String
' Récupérer le nom d'un fichier sans son extention

Dim i As Long

      If Len(filename) > 0 Then  ' Si la chaine n'est pas vide
      
            ' Passe tous les caractères un à un à partir de la fin
            For i = Len(filename) To 1 Step -1
            
                  ' Si on est arriver sur un point(.) on quitte la boucle
                  If Mid$(filename, i, 1) = "." Then Exit For
                  
            Next i
            
            ' Renvoie le nom du fichier sans son extention
            RemoveExtention = Left$(filename, i - 1)
            
      Else
      
            ' Sinon on renvoie rien
            RemoveExtention = ""
            
      End If

End Function

Function PathExists(ByVal FullPath As String) As Boolean
' Vérifier l'existance d'un dossier ou répertoire

      If Len(FullPath) > 0 Then  ' Si la chaine n'est pas vide(Null)
      
            ' Renvoie le résultat de la recherche
            PathExists = CBool(Len(Dir$(FullPath, vbDirectory)))
            
      End If

End Function

Function FileExists(ByVal filename As String) As Boolean
' Vérifier l'existance d'un fichier

      If Len(filename) > 0 Then  ' Si la chaine n'est pas vide(Null)
      
            ' Renvoie le résultat de la recherche
            FileExists = CBool(Len(Dir$(filename, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)))
            
      End If

End Function
