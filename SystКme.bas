Attribute VB_Name = "mdMisc"
'Winux Graphic User Interface for Windows based systems
'Copyright (C) 2002-2003 Winux Team
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'A copy of this licence is available in root\system directory.
'http://www.winux.free.fr or tex_winux@hotmail.com for more details.

'Déclaration pour identification noyau NT
Public isNT As Boolean

'Déclaration pour l'identification utilisateur
Public login As String

'Declarations pour la RAM
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Dim MS As MEMORYSTATUS

'********************************************************
'CPU
'
'mdlMisc - copyright © 2001, The KPD-Team
'Visit our site at http://www.allapi.net
'or email us at KPDTeam@allapi.net
Option Explicit
Dim Cnt1 As Long 'variable de comptage, incrémentation
Dim Cnt2 As Long 'variable de comptage, incrémentation
Const SPACE = 5
Const BAR_WIDTH = 50
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const HIGH_PRIORITY_CLASS = &H80
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private GraphPoints(0 To 99) As Long
Sub DrawUsage(lUsage As Long, picPercent As PictureBox, picGraph As PictureBox)
    picPercent.ScaleMode = vbPixels
    For Cnt1 = 0 To 10
        picPercent.Line (SPACE, SPACE + Cnt1 * 3)-(SPACE + BAR_WIDTH, SPACE + Cnt1 * 3 + 1), IIf(lUsage >= 100 - Cnt1 * 10 And lUsage <> 0, &HC000&, &H4000&), BF
    Next Cnt1
    ShiftPoints
    GraphPoints(UBound(GraphPoints)) = lUsage
    picGraph.Cls
    For Cnt1 = LBound(GraphPoints) To UBound(GraphPoints) - 1
        picGraph.Line (Cnt1, 100 - GraphPoints(Cnt1))-(Cnt1 + 1, 100 - GraphPoints(Cnt1 + 1)), &H8000&
    Next Cnt1
End Sub
'Shift all the points from the graph one place to the left
Sub ShiftPoints()
    For Cnt2 = LBound(GraphPoints) To UBound(GraphPoints) - 1
        GraphPoints(Cnt2) = GraphPoints(Cnt2 + 1)
    Next Cnt2
End Sub
'return True is the OS is WindowsNT3.5(1), NT4.0, 2000 or XP
Public Function IsWinNT() As Boolean
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'retrieve OS version info
    GetVersionEx OSInfo
    'if we're on NT, return True
    IsWinNT = (OSInfo.dwPlatformId = 2)
End Function

'********************************************************
'RAM
'code réalisé à partir d'une source de FX
'Merci à son auteur de l'avoir fait partagé sur vbfrance.

'Pour obtenir la memoire total
Function TotalMemory(Mo As Boolean)
MS.dwLength = Len(MS)
GlobalMemoryStatus MS
If Mo = True Then
TotalMemory = Int(MS.dwTotalPhys / 1024 / 1024) & " Mo"
Exit Function
Else
TotalMemory = MS.dwTotalPhys
End If
End Function

'Pour obtenir la memoire libre
Function GetFreeMemory()
MS.dwLength = Len(MS)
GlobalMemoryStatus MS
GetFreeMemory = MS.dwAvailPhys
End Function
