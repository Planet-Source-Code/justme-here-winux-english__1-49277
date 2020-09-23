VERSION 5.00
Begin VB.Form Bureau 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Office"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer DateEtTemps 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7680
      Top             =   120
   End
   Begin VB.PictureBox picUsage 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   9210
      ScaleHeight     =   61
      ScaleMode       =   0  'User
      ScaleWidth      =   62
      TabIndex        =   1
      Top             =   0
      Width           =   975
      Begin VB.Label lblCpuUsage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   263
         Left            =   0
         TabIndex        =   2
         Top             =   620
         Width           =   945
      End
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10320
      Top             =   240
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000002&
      Height          =   975
      Left            =   10185
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Line Line23 
      X1              =   73
      X2              =   73
      Y1              =   15
      Y2              =   48
   End
   Begin VB.Line Line22 
      BorderColor     =   &H8000000F&
      X1              =   40
      X2              =   40
      Y1              =   15
      Y2              =   48
   End
   Begin VB.Line Line21 
      X1              =   141
      X2              =   141
      Y1              =   15
      Y2              =   48
   End
   Begin VB.Line Line20 
      BorderColor     =   &H8000000F&
      X1              =   109
      X2              =   109
      Y1              =   16
      Y2              =   48
   End
   Begin VB.Line Line19 
      X1              =   209
      X2              =   209
      Y1              =   15
      Y2              =   48
   End
   Begin VB.Line Line18 
      BorderColor     =   &H8000000F&
      X1              =   177
      X2              =   177
      Y1              =   16
      Y2              =   48
   End
   Begin VB.Line Line17 
      X1              =   277
      X2              =   277
      Y1              =   15
      Y2              =   48
   End
   Begin VB.Line Line16 
      X1              =   245
      X2              =   278
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line Line15 
      X1              =   177
      X2              =   210
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line Line14 
      X1              =   109
      X2              =   142
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line Line13 
      X1              =   40
      X2              =   74
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000F&
      X1              =   41
      X2              =   73
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line11 
      BorderColor     =   &H8000000F&
      X1              =   109
      X2              =   141
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000F&
      X1              =   177
      X2              =   209
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000F&
      X1              =   245
      X2              =   277
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000F&
      X1              =   245
      X2              =   245
      Y1              =   16
      Y2              =   48
   End
   Begin VB.Line Line7 
      X1              =   312
      X2              =   344
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line Line6 
      X1              =   344
      X2              =   344
      Y1              =   15
      Y2              =   49
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000F&
      X1              =   312
      X2              =   344
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000F&
      X1              =   312
      X2              =   312
      Y1              =   16
      Y2              =   48
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   600
      Picture         =   "Bureau.frx":0000
      ToolTipText     =   "Trash can"
      Top             =   240
      Width           =   510
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1680
      Picture         =   "Bureau.frx":0442
      Stretch         =   -1  'True
      ToolTipText     =   "Search for..."
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2640
      Picture         =   "Bureau.frx":0884
      Stretch         =   -1  'True
      ToolTipText     =   "Office"
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3720
      Picture         =   "Bureau.frx":0CC6
      Stretch         =   -1  'True
      ToolTipText     =   "Console"
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4680
      Picture         =   "Bureau.frx":1108
      Stretch         =   -1  'True
      ToolTipText     =   "Shut Down"
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   8085
      Left            =   0
      Stretch         =   -1  'True
      Top             =   975
      Width           =   12000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000F&
      X1              =   381
      X2              =   381
      Y1              =   0
      Y2              =   63
   End
   Begin VB.Label lblTemps 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   6480
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblMemDispo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   8400
      TabIndex        =   4
      Top             =   630
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RAM"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   8415
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   8235
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   548
      X2              =   548
      Y1              =   0
      Y2              =   64
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   0
      X2              =   800
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000016&
      FillColor       =   &H80000016&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Bureau"
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

Option Explicit
Private QueryObject As Object  'Pour l'usage du processeur

Private Sub DateEtTemps_Timer()
    'Affichage date et heure
    If (lblTemps.Caption <> FormatDateTime(Time, 4)) Then
        lblTemps.Caption = FormatDateTime(Time, 4)
        lblDate.Caption = Format$(Date, "dddd d mmmm yyyy")
    End If
End Sub

Private Sub Form_Activate()
    'on place le bureau en arrière plan
    Call SetWindowPos(Bureau.hWnd, 1, 0, 0, 0, 0, &H2 Or &H1 Or &H40 Or &H10)
End Sub

Private Sub Form_Load()
    'on place le bureau en arrière plan
    Call SetWindowPos(Bureau.hWnd, 1, 0, 0, 0, 0, &H2 Or &H1 Or &H40 Or &H10)
    'on cache la taskbar
    Dim rtn As Long
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    'on place tout de suite le formulaire correctement
    Resolution
    'donne une priorité haute au process de l'appli (feuille?)
    'ainsi on est sûr que l'appli est opérationelle même
    'si un autre process consomme beaucoup de cycles CPU
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    'initialisation QueryOject
    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
        isNT = True 'noyau NT ou pas
    Else
        Set QueryObject = New clsCPUUsage
        isNT = False 'noyau NT ou pas
    End If
    'Initialisation nécessaire pour recevoir des valeurs correctes
    QueryObject.Initialize
    'Initialise le Timers
    tmrRefresh.Enabled = True
    DateEtTemps.Enabled = True
    'on n'attend pas le premier Interval des Timer
    tmrRefresh_Timer
    DateEtTemps_Timer
    'on initialise les paramètres de l'utilisateur
    InitUser
    'on initialise le témoin de compression/décompression
    DCencour = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu Menus.buro
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Stop les Timers
    tmrRefresh.Enabled = False
    DateEtTemps.Enabled = False
    'déchargement QueryObject
    QueryObject.Terminate
    Set QueryObject = Nothing
    'faire réapparaitre la taskbar
    Dim rtn As Long
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub

Private Sub Image2_Click()
Arrêt.Show 1, Bureau
End Sub

Private Sub Image3_Click()
Home.Show 0, Bureau
End Sub

Private Sub Image4_Click()
Explorateur.Show 0, Bureau
End Sub

Private Sub Image5_Click()
Cherche.Show 0, Bureau
End Sub

Private Sub Image6_Click()
Explorateur.Dir1.Path = App.Path & "\home\" & login & "\Poubelle"
Explorateur.Show 0, Bureau
End Sub

Private Sub tmrRefresh_Timer()
    Dim Ret As Long
    'Récupératon de l'usage du CPU
    Ret = QueryObject.Query
    If Ret = -1 Then
        tmrRefresh.Enabled = False
        lblCpuUsage.Caption = ":("
        MsgBox "Error while retrieving CPU usage"
    Else
        DrawUsage Ret, picUsage, picGraph
        lblCpuUsage.Caption = CStr(Ret) + "%"
    End If
    'Récupération de la ram disponible
    lblMemDispo = Int(GetFreeMemory / 1024)
End Sub

'adaptation du Bureau à la résolution de l'écran
'by tex 25/10/02 20:03:12(GMT)
'status : ok
'utilisation :
'   Resolution
Private Sub Resolution()
    'on replace les controls suivant la résolution d'écran
    Me.Height = Screen.Height
    Image1.Height = ((Me.Height / 15) - Image1.Top)
    Me.Width = Screen.Width
    Image1.Width = (Me.Width / 15)
    Shape1.Width = (Me.Width / 15)
    Dim taket As Long
    taket = Shape1.Width - (picGraph.Left + picGraph.Width)
    picGraph.Left = picGraph.Left + taket
    picUsage.Left = picUsage.Left + taket
    Shape2.Left = Shape2.Left + taket
    Label1.Left = Label1.Left + taket
    lblMemDispo.Left = lblMemDispo.Left + taket
    Line4.X1 = Line4.X1 + taket
    Line4.X2 = Line4.X2 + taket
    lblTemps.Left = lblTemps.Left + taket
    lblDate.Left = lblDate.Left + taket
    Line2.X1 = Shape2.Left
    Line2.X2 = Shape2.Left
End Sub
