VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHuffman"
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

Private Const PROGRESS_CALCFREQUENCY = 7
Private Const PROGRESS_CALCCRC = 5
Private Const PROGRESS_ENCODING = 88
Private Const PROGRESS_DECODING = 89
Private Const PROGRESS_CHECKCRC = 11
Event Progress(Procent As Integer)

Private Type HUFFMANTREE
  ParentNode As Integer
  RightNode As Integer
  LeftNode As Integer
  Value As Integer
  Weight As Long
End Type

Private Type ByteArray
  Count As Byte
  Data() As Byte
End Type

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub DecodeFile(SourceFile As String, DestFile As String)

  Dim ByteArray() As Byte
  Dim Filenr As Integer
  
  If (Not FileExist(SourceFile)) Then
    Err.Raise vbObjectError, "clsHuffman.DecodeFile()", "Source file does not exist"
  End If
  
  Filenr = FreeFile
  Open SourceFile For Binary As #Filenr
  ReDim ByteArray(0 To LOF(Filenr) - 1)
  Get #Filenr, , ByteArray()
  Close #Filenr
  
  Call DecodeByte(ByteArray(), UBound(ByteArray) + 1)
  If (FileExist(DestFile)) Then Kill DestFile
  
  Open DestFile For Binary As #Filenr
  Put #Filenr, , ByteArray()
  Close #Filenr

End Sub
Private Sub CreateTree(Nodes() As HUFFMANTREE, NodesCount As Long, Char As Long, Bytes As ByteArray)

  Dim a As Integer
  Dim NodeIndex As Long
  
  NodeIndex = 0
  For a = 0 To (Bytes.Count - 1)
    If (Bytes.Data(a) = 0) Then
      If (Nodes(NodeIndex).LeftNode = -1) Then
        Nodes(NodeIndex).LeftNode = NodesCount
        Nodes(NodesCount).ParentNode = NodeIndex
        Nodes(NodesCount).LeftNode = -1
        Nodes(NodesCount).RightNode = -1
        Nodes(NodesCount).Value = -1
        NodesCount = NodesCount + 1
      End If
      NodeIndex = Nodes(NodeIndex).LeftNode
    ElseIf (Bytes.Data(a) = 1) Then
      If (Nodes(NodeIndex).RightNode = -1) Then
        Nodes(NodeIndex).RightNode = NodesCount
        Nodes(NodesCount).ParentNode = NodeIndex
        Nodes(NodesCount).LeftNode = -1
        Nodes(NodesCount).RightNode = -1
        Nodes(NodesCount).Value = -1
        NodesCount = NodesCount + 1
      End If
      NodeIndex = Nodes(NodeIndex).RightNode
    Else
      Stop
    End If
  Next
  
  Nodes(NodeIndex).Value = Char

End Sub
Public Function DecodeString(Text As String) As String
  
  Dim ByteArray() As Byte
  
  ByteArray() = StrConv(Text, vbFromUnicode)
  
  Call DecodeByte(ByteArray, Len(Text))
  
  DecodeString = StrConv(ByteArray(), vbUnicode)
  
End Function

Public Sub DecodeByte(ByteArray() As Byte, ByteLen As Long)
  
  Dim i As Long
  Dim j As Long
  Dim Pos As Long
  Dim Char As Byte
  Dim CurrPos As Long
  Dim Count As Integer
  Dim CheckSum As Byte
  Dim Result() As Byte
  Dim BitPos As Integer
  Dim NodeIndex As Long
  Dim ByteValue As Byte
  Dim ResultLen As Long
  Dim NodesCount As Long
  Dim lResultLen As Long
  Dim NewProgress As Integer
  Dim CurrProgress As Integer
  Dim BitValue(0 To 7) As Byte
  Dim Nodes(0 To 511) As HUFFMANTREE
  Dim CharValue(0 To 255) As ByteArray
  
  If (ByteArray(0) <> 72) Or (ByteArray(1) <> 69) Or (ByteArray(3) <> 13) Then
  ElseIf (ByteArray(2) = 48) Then
    'Decode = Mid$(Text, 5)
    Call CopyMem(ByteArray(0), ByteArray(4), ByteLen - 4)
    ReDim Preserve ByteArray(0 To ByteLen - 5)
    Exit Sub
  ElseIf (ByteArray(2) <> 51) Then
    Err.Raise vbObjectError, "HuffmanDecode()", "The data either was not compressed with HE3 or is corrupt (identification string not found)"
    Exit Sub
  End If
  
  CurrPos = 5
    
  CheckSum = ByteArray(CurrPos - 1)
  CurrPos = CurrPos + 1
  
  Call CopyMem(ResultLen, ByteArray(CurrPos - 1), 4)
  CurrPos = CurrPos + 4
  lResultLen = ResultLen
  
  If (ResultLen = 0) Then Exit Sub
  
  ReDim Result(0 To ResultLen - 1)
  
  Call CopyMem(Count, ByteArray(CurrPos - 1), 2)
  CurrPos = CurrPos + 2
  
  For i = 1 To Count
    With CharValue(ByteArray(CurrPos - 1))
      CurrPos = CurrPos + 1
      .Count = ByteArray(CurrPos - 1)
      CurrPos = CurrPos + 1
      ReDim .Data(0 To .Count - 1)
    End With
  Next
  
  For i = 0 To 7
    BitValue(i) = 2 ^ i
  Next
  
  ByteValue = ByteArray(CurrPos - 1)
  CurrPos = CurrPos + 1
  BitPos = 0
  For i = 0 To 255
    With CharValue(i)
      If (.Count > 0) Then
        For j = 0 To (.Count - 1)
          If (ByteValue And BitValue(BitPos)) Then .Data(j) = 1
          BitPos = BitPos + 1
          If (BitPos = 8) Then
            ByteValue = ByteArray(CurrPos - 1)
            CurrPos = CurrPos + 1
            BitPos = 0
          End If
        Next
      End If
    End With
  Next
  If (BitPos = 0) Then CurrPos = CurrPos - 1
  
  NodesCount = 1
  Nodes(0).LeftNode = -1
  Nodes(0).RightNode = -1
  Nodes(0).ParentNode = -1
  Nodes(0).Value = -1
  For i = 0 To 255
    Call CreateTree(Nodes(), NodesCount, i, CharValue(i))
  Next
  
  ResultLen = 0
  For CurrPos = CurrPos To ByteLen
    ByteValue = ByteArray(CurrPos - 1)
    For BitPos = 0 To 7
      If (ByteValue And BitValue(BitPos)) Then
        NodeIndex = Nodes(NodeIndex).RightNode
      Else
        NodeIndex = Nodes(NodeIndex).LeftNode
      End If
      If (Nodes(NodeIndex).Value > -1) Then
        Result(ResultLen) = Nodes(NodeIndex).Value
        ResultLen = ResultLen + 1
        If (ResultLen = lResultLen) Then GoTo DecodeFinished
        NodeIndex = 0
      End If
    Next
    If (CurrPos Mod 10000 = 0) Then
      NewProgress = CurrPos / ByteLen * PROGRESS_DECODING
      If (NewProgress <> CurrProgress) Then
        CurrProgress = NewProgress
        RaiseEvent Progress(CurrProgress)
      End If
    End If
  Next
DecodeFinished:

  Char = 0
  For i = 0 To (ResultLen - 1)
    Char = Char Xor Result(i)
    If (i Mod 10000 = 0) Then
      NewProgress = i / ResultLen * PROGRESS_CHECKCRC + PROGRESS_DECODING
      If (NewProgress <> CurrProgress) Then
        CurrProgress = NewProgress
        RaiseEvent Progress(CurrProgress)
      End If
    End If
  Next
  If (Char <> CheckSum) Then
    Err.Raise vbObjectError, "clsHuffman.Decode()", "The data might be corrupted (checksum did not match expected value)"
  End If

  ReDim ByteArray(0 To ResultLen - 1)
  Call CopyMem(ByteArray(0), Result(0), ResultLen)
  
  If (CurrProgress <> 100) Then
    RaiseEvent Progress(100)
  End If
  
End Sub

Private Function FileExist(filename As String) As Boolean

  On Error GoTo FileDoesNotExist
  
  Call FileLen(filename)
  FileExist = True
  Exit Function
  
FileDoesNotExist:
  FileExist = False
  
End Function


