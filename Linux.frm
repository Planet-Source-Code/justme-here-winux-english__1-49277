VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Konsole 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Konsole"
   ClientHeight    =   3645
   ClientLeft      =   1305
   ClientTop       =   2130
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   660
      Top             =   1740
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2280
      Top             =   2040
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Linux 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   3645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5385
   End
End
Attribute VB_Name = "Konsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Le code de cette form a été réalisé à partir d'une source de Xentor
'Merci à son auteur de l'avoir faite partagée sur vbfrance.
Option Explicit
Dim afficheTiret As Boolean
Dim texte As String
Dim UserText As String
Dim groscaracteres As Boolean
Dim invite As String
Dim curPath As String
Dim fs As FileSystemObject
Dim phase As Integer
Dim fold As Folder
Dim password As String
Dim HaveToWait As Boolean
Dim xfiles As File
Dim xDrives As Drive
Dim xfold As Folder
Dim newLogin As String
Dim newPass As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 1 Then
groscaracteres = True
End If
Select Case KeyCode
Case 8 'retour arrière
If Len(UserText) > 0 Then UserText = Left(UserText, Len(UserText) - 1)
Case 20 'majuscule
groscaracteres = Not groscaracteres
Case 13 'entrée
    If HaveToWait = False Then
        texte = texte & invite & UserText & Chr(13)
        analyse
    Else
        Select Case phase
            Case 1:
                login = UserText
                texte = texte & login
                echo ""
                echo "password : ", False
                phase = 2
            Case 2:
                password = UserText
                echo ""
                    HaveToWait = False
                    On Error GoTo 1
                    If login = "root" Then
                        Open (App.Path & login & "\" & login) For Input As #1
                    Else
                        Open (App.Path & "\home\" & login & "\" & login) For Input As #1
                    End If
                    Input #1, cod
                    Close #1
                    cod = Cryptage(cod, 2002, 1)
                    If cod = password Then
                        If KonsoleMod = False Then InitUser 'on initialise les paramètres de l'utilisateur
                        changeInvite
                        Bureau.Show
                        Unload Me
                    Else
                        echo "Pass ou login invalide !"
                        logout
                    End If
                    Exit Sub
1:                     echo "Pass ou login invalide !"
                    logout
            Case 6:
                If UserText = "." Then
                    HaveToWait = False
                    WS.close
                Else

                    WS.SendData UserText & vbCrLf
                    texte = texte & UserText
                    echo ""
                End If
            End Select
        End If
        UserText = ""
        Case 110 'point
        UserText = UserText & "."
        Case 191 '2 points
        UserText = UserText & ":"
Case Else
If KeyCode >= 96 And KeyCode <= 105 Then
UserText = UserText & (KeyCode - 96)
ElseIf KeyCode >= 48 And KeyCode <= 57 And groscaracteres = False Then
UserText = UserText & Mid("à&é""'(-è_ç", (KeyCode - 48) + 1, 1)
ElseIf KeyCode = 188 And groscaracteres = False Then UserText = UserText & ","
ElseIf KeyCode = 188 And groscaracteres Then UserText = UserText & "?"
ElseIf KeyCode = 190 And groscaracteres Then UserText = UserText & "."
ElseIf KeyCode = 56 And groscaracteres Then UserText = UserText & "\"
ElseIf KeyCode = 16 Then 'pas de caractère quand on appuie sur maj
ElseIf KeyCode = 37 Then
ElseIf KeyCode = 38 Then
ElseIf KeyCode = 39 Then
ElseIf KeyCode = 40 Then
Else
UserText = UserText & IIf(groscaracteres, Chr(KeyCode), LCase(Chr(KeyCode)))
End If
End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 16 Then groscaracteres = False
End Sub
Sub analyse()
Dim commandExists As Boolean
Dim param As String
Dim i As Integer
commandExists = True
Select Case LCase(UserText)
    Case "man"
        echo "Commandes acceptées : man, cd, ls, bye, clear, mkdir, rmdir, cat, rm, wd, telnet, quit, ld, cd .., exit"
    Case "clear"
        texte = ""
    Case "bye"
        logout
    Case "exit"
        Unload Me
    Case "quit"
        Unload Arrêt
        Unload Explorateur
        Unload Menus
        Unload Me
        Unload Bureau
        End
    Case "wd"
        echo curPath
    Case "ls"
        If curPath <> "\" Then
            If Len(curPath) = 1 Then curPath = curPath & ":\"
            Set fold = fs.GetFolder(curPath)
            For Each xfold In fold.SubFolders
                echo xfold.Name & "\"
            Next
            For Each xfiles In fold.Files
                echo xfiles.Name
            Next
        Else
            For Each xDrives In fs.Drives
                echo xDrives.DriveLetter & "\"
            Next
        End If
    Case "cd .."
        curPath = Left(curPath, (Len(curPath) - Len(dossier)))
    Case Else
        commandExists = False
End Select
If LCase(Left(UserText, 2)) = "cd" Then
    commandExists = True
    param = LTrim(RTrim(Right(UserText, Len(UserText) - 2)))
        If curPath = "\" Then
            Dim tempLect As String
            tempLect = param & ":\"
            If fs.DriveExists(tempLect) Then
                curPath = UCase(param)
            Else
                echo "Disque inexistant. Tapez ls pour avoir la liste des lecteurs disponibles."
            End If
        Else
            If Right(param, 1) = "\" Then param = Left(param, Len(param) - 1)
            Dim newFolder As String
            newFolder = curPath & "\" & param
            If fs.FolderExists(newFolder) Then Set fold = fs.GetFolder(newFolder)
            If fold.IsRootFolder Then
                curPath = Left(fold.Path, 1)
            ElseIf Len(curPath) = 1 Then
                curPath = "\"
            Else
                curPath = fold.Path
            End If
        End If
    ElseIf LCase(Left(UserText, 5)) = "mkdir" Then
        commandExists = True
        If curPath <> "\" Then
            param = LTrim(RTrim(Right(UserText, Len(UserText) - 5)))
            If Not fs.FolderExists(curPath & "\" & param) Then
                fs.CreateFolder IIf(Len(curPath) = 1, curPath & ":", curPath) & "\" & param
            Else
                echo "Ce dossier existe déjà."
            End If
        Else
            echo "Opération non valide à ce stade."
        End If
    ElseIf LCase(Left(UserText, 2)) = "ld" Then
        commandExists = True
        param = LTrim(RTrim(Right(UserText, Len(UserText) - 2)))
        If fs.FileExists(curPath & "\" & param) Then
         Shell param, vbNormalFocus
        Else
            echo "Ce fichier n'existe pas."
        End If
    ElseIf LCase(Left(UserText, 5)) = "rmdir" Then
        commandExists = True
        If curPath <> "\" Then
            param = LTrim(RTrim(Right(UserText, Len(UserText) - 5)))
            If fs.FolderExists(curPath & "\" & param) Then
                fs.DeleteFolder IIf(Len(curPath) = 1, curPath & ":", curPath) & "\" & param
            Else
                echo "Ce dossier n'existe pas."
            End If
        Else
            echo "Opération non valide à ce stade."
        End If
    ElseIf LCase(Left(UserText, 2)) = "rm" Then
        commandExists = True
        If curPath <> "\" Then
            param = LTrim(RTrim(Right(UserText, Len(UserText) - 2)))
            If fs.FileExists(curPath & "\" & param) Then
                fs.DeleteFile IIf(Len(curPath) = 1, curPath & ":", curPath) & "\" & param
            Else
                echo "Ce fichier n'existe pas."
            End If
        Else
            echo "Opération non valide à ce stade."
        End If
    ElseIf LCase(Left(UserText, 4)) = "cat" Then
        commandExists = True
        If curPath <> "\" Then
            param = LTrim(RTrim(Right(UserText, Len(UserText) - 4)))
            If fs.FileExists(curPath & "\" & param) Then
                texte = texte & fs.OpenTextFile(curPath & "\" & param).ReadAll
                echo ""
            Else
                echo "Ce fichier n'existe pas."
            End If
        Else
            echo "Opération non valide à ce stade."
        End If
    ElseIf LCase(Left(UserText, 6)) = "telnet" Then
        commandExists = True
        Dim server As String, port As String
        param = LTrim(RTrim(Right(UserText, Len(UserText) - 6)))
        server = Left(param, InStr(1, param, " ") - 1)
        port = Right(param, Len(param) - InStrRev(param, " "))
        Select Case LCase(port)
        Case "http":
        port = 80
        Case "pop3":
        port = 110
        Case "smtp":
        port = 125
        Case "irc":
        port = 6667
        Case "ftp":
        port = 25
        End Select
        If Not IsNumeric(port) Then echo "Le port doit être une valeur numérique !": Exit Sub
        WS.RemoteHost = server
        WS.RemotePort = port
WS.close
DoEvents
WS.Connect
DoEvents
        HaveToWait = True
        phase = 6
End If
changeInvite
If commandExists = False Then echo "Commande non reconnue. Tapez man pour avoir la liste des commandes disponibles."
End Sub
Private Sub Form_Load()
Me.Show
afficheTiret = True
curPath = App.Path
Set fs = New FileSystemObject
If KonsoleMod = False Then
    logout
Else
    curPath = curPath & "\" & login
    texte = ""
    HaveToWait = False
End If
End Sub
Sub changeInvite()
If login = "root" Then
    invite = "[" & login & "@localhost " & dossier & "/]# "
Else
    invite = "[" & login & "@localhost " & dossier & "/]$ "
End If
If KonsoleMod = False Then
    Bureau.Show
    Unload Me
End If
End Sub
Sub echo(str As String, Optional AlaLigne = True)
texte = texte & str & IIf(AlaLigne, Chr(13), "")
actualiser
End Sub
Function dossier() As String
If curPath = "\" Then dossier = "\" Else dossier = Right(curPath, Len(curPath) - InStrRev(curPath, "\"))
End Function

Private Sub Form_Unload(Cancel As Integer)
KonsoleMod = False
End Sub

Private Sub Timer1_Timer()
afficheTiret = Not afficheTiret
actualiser
End Sub
Sub actualiser()
Dim added As String
If HaveToWait = False Then added = invite & UserText & IIf(afficheTiret, "_", "") Else added = UserText & IIf(afficheTiret, "_", "")
Dim nbRetours As Integer, i As Integer
Dim caracts As Integer
For i = Len(texte) To 1 Step -1
If Asc(Mid(texte, i, 1)) = "13" Or caracts > 44 Then
nbRetours = nbRetours + 1
caracts = 0
If nbRetours >= 14 Then texte = Right(texte, Len(texte) - i): Exit For
Else
caracts = caracts + 1
End If
Next
Debug.Print nbRetours
Linux.Caption = texte & added
End Sub

Private Sub Timer2_Timer()
actualiser
End Sub
Sub logout()
HaveToWait = True
echo "localhost login : ", False
phase = 1
End Sub

Private Sub WS_Connect()
echo "Connecté à " & WS.RemoteHost & "."
echo "Terminez votre session avec le caractère ""."""
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
Dim vbstring As String
WS.GetData vbstring, 1
echo vbstring, False
End Sub

Private Sub WS_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
WS.close
echo "Erreur " & number & " : " & Description
HaveToWait = False
End Sub

