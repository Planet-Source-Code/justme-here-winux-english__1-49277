Attribute VB_Name = "Compression"
'Winux Graphic User Interface for Windows based systems
'Copyright (C) 2002-2003 Winux Team
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'A copy of this licence is available in root\system directory.
'http://www.winux.free.fr or tex_winux@hotmail.com for more details.

'variable pour opération D/C en cours ou libre
Public DCencour As Boolean

'pour utiliser ce module
'mettre en général de la feuille:
'Dim PIn As String, POut As String, LevelC As Integer, Max As Long, Verif As String
'
'Pour compresser:
'LevelC = 9
'PIn = File1.Path & "\" & File1.FileName
'POut = PIn & "_"
'result = Compression.CompressFile(PIn, POut, LevelC)
'
'Pour décompresser:
'PIn = File1.Path & "\" & File1.FileName
'Verif = Right(PIn, 1)
'If Verif <> "_" Then
'MsgBox ("Ce fichier n'est pas un fichier compressé au format Softzip !")
'GoTo 1
'End If
'Max = Len(PIn) - 1
'POut = Mid(PIn, 1, Max)
'result = Compression.DecompressFile(PIn, POut)

Option Explicit
'Déclaration
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'Variable
Dim lngCompressedSize As Long
Dim lngDecompressedSize As Long
Enum CZErrors 'Constante pour la compression/décompression
    Z_OK = 0
    Z_STREAM_END = 1
    Z_NEED_DICT = 2
    Z_ERRNO = -1
    Z_STREAM_ERROR = -2
    Z_DATA_ERROR = -3
    Z_MEM_ERROR = -4
    Z_BUF_ERROR = -5
    Z_VERSION_ERROR = -6
End Enum

Enum CompressionLevels 'Constante pour la compression/décompression
    Z_NO_COMPRESSION = 0
    Z_BEST_SPEED = 1
    'Les Levels 2-8 existe aussi
    Z_BEST_COMPRESSION = 9
    Z_DEFAULT_COMPRESSION = -1
End Enum

Private Property Get ValueCompressedSize() As Long '
    'Taille de l'objet après compression
    ValueCompressedSize = lngCompressedSize
End Property

Private Property Let ValueCompressedSize(ByVal New_ValueCompressedSize As Long)
    lngCompressedSize = New_ValueCompressedSize
End Property

Private Property Get ValueDecompressedSize() As Long '
    'Taille de l'objet après la décompression
    ValueDecompressedSize = lngDecompressedSize
End Property

Private Property Let ValueDecompressedSize(ByVal New_ValueDecompressedSize As Long)
    lngDecompressedSize = New_ValueDecompressedSize
End Property

Private Function CompressByteArray(TheData() As Byte, CompressionLevel As Integer) As Long '
    'compression par blocs d'octets
    Dim lngResult As Long
    Dim lngBufferSize As Long
    Dim arrByteArray() As Byte
    lngDecompressedSize = UBound(TheData) + 1
    'alloué de la mémoire pour les blocs d'octets
    lngBufferSize = UBound(TheData) + 1
    lngBufferSize = lngBufferSize + (lngBufferSize * 0.01) + 12
    ReDim arrByteArray(lngBufferSize)
    'compression par blocs d'octets des données
    lngResult = compress2(arrByteArray(0), lngBufferSize, TheData(0), UBound(TheData) + 1, CompressionLevel)
    'tronquer pour obtenir la taille compressée
    ReDim Preserve TheData(lngBufferSize - 1)
    CopyMemory TheData(0), arrByteArray(0), lngBufferSize
    'définition de la propriété
    lngCompressedSize = UBound(TheData) + 1
    'retourner le code d'erreur si erreur il y a
    CompressByteArray = lngResult
End Function

Private Function CompressString(Text As String, CompressionLevel As Integer) As Long '
    'compression d'une chaîne
    Dim lngOrgSize As Long
    Dim lngReturnValue As Long
    Dim lngCmpSize As Long
    Dim strTBuff As String
    ValueDecompressedSize = Len(Text)
    'allocation de chaînes d'espace pour les tampons
    lngOrgSize = Len(Text)
    strTBuff = String(lngOrgSize + (lngOrgSize * 0.01) + 12, 0)
    lngCmpSize = Len(strTBuff)
    'compression des chaînes de donnée (chaîne temporaires dans les tampons)
    lngReturnValue = compress2(ByVal strTBuff, lngCmpSize, ByVal Text, Len(Text), CompressionLevel)
    'Crop the string and set it to the actual string.
    Text = Left$(strTBuff, lngCmpSize)
    'Set compressed size of string.
    ValueCompressedSize = lngCmpSize
    'Cleanup
    strTBuff = ""
    'return error code (if any)
    CompressString = lngReturnValue
End Function

Private Function DecompressByteArray(TheData() As Byte, OriginalSize As Long) As Long '
    'decompress a byte array
    Dim lngResult As Long
    Dim lngBufferSize As Long
    Dim arrByteArray() As Byte
    lngDecompressedSize = OriginalSize
    lngCompressedSize = UBound(TheData) + 1
    'Allocate memory for byte array
    lngBufferSize = OriginalSize
    lngBufferSize = lngBufferSize + (lngBufferSize * 0.01) + 12
    ReDim arrByteArray(lngBufferSize)
    'Decompress data
    lngResult = uncompress(arrByteArray(0), lngBufferSize, TheData(0), UBound(TheData) + 1)
    'Truncate buffer to compressed size
    ReDim Preserve TheData(lngBufferSize - 1)
    CopyMemory TheData(0), arrByteArray(0), lngBufferSize
    'return error code (if any)
    DecompressByteArray = lngResult
End Function

Private Function DecompressString(Text As String, OriginalSize As Long) As Long '
    'decompress a string
    Dim lngResult As Long
    Dim lngCmpSize As Long
    Dim strTBuff As String
    'Allocate string space
    strTBuff = String(ValueDecompressedSize + (ValueDecompressedSize * 0.01) + 12, 0)
    lngCmpSize = Len(strTBuff)
    ValueDecompressedSize = OriginalSize
    'Decompress
    lngResult = uncompress(ByVal strTBuff, lngCmpSize, ByVal Text, Len(Text))
    'Make string the size of the uncompressed string
    Text = Left$(strTBuff, lngCmpSize)
    ValueCompressedSize = lngCmpSize
    'return error code (if any)
    DecompressString = lngResult
End Function
Public Function CompressFile(FilePathIn As String, FilePathOut As String, CompressionLevel As Integer) As Long
    'compress a file
    Dim intNextFreeFile As Integer
    Dim TheBytes() As Byte
    Dim lngResult As Long
    Dim lngFileLen As Long
    lngFileLen = FileLen(FilePathIn)
    'allocate byte array
    ReDim TheBytes(lngFileLen - 1)
    'read byte array from file
    Close #10
    intNextFreeFile = FreeFile '10 'FreeFile
    Open FilePathIn For Binary Access Read As #intNextFreeFile
        Get #intNextFreeFile, , TheBytes()
    Close #intNextFreeFile
    'compress byte array
    lngResult = CompressByteArray(TheBytes(), CompressionLevel)
    'kill any file in place
    On Error Resume Next
    Kill FilePathOut
    On Error GoTo 0
    'Write it out
    intNextFreeFile = FreeFile
    Open FilePathOut For Binary Access Write As #intNextFreeFile
        Put #intNextFreeFile, , lngFileLen 'must store the length of the original file
        Put #intNextFreeFile, , TheBytes()
    Close #intNextFreeFile
    Erase TheBytes
    CompressFile = lngResult
    DCencour = False
End Function

Public Function DecompressFile(FilePathIn As String, FilePathOut As String) As Long
    'decompress a file
    Dim intNextFreeFile As Integer
    Dim TheBytes() As Byte
    Dim lngResult As Long
    Dim lngFileLen As Long
    'allocate byte array
    ReDim TheBytes(FileLen(FilePathIn) - 1)
    'read byte array from file
    intNextFreeFile = FreeFile
    Open FilePathIn For Binary Access Read As #intNextFreeFile
        Get #intNextFreeFile, , lngFileLen 'the original (uncompressed) file's length
        Get #intNextFreeFile, , TheBytes()
    Close #intNextFreeFile
    'decompress
    lngResult = DecompressByteArray(TheBytes(), lngFileLen)
    'kill any file already there
    On Error Resume Next
    Kill FilePathOut
    On Error GoTo 0
    'Write it out
    intNextFreeFile = FreeFile
    Open FilePathOut For Binary Access Write As #intNextFreeFile
        Put #intNextFreeFile, , TheBytes()
    Close #intNextFreeFile
    Erase TheBytes
    DecompressFile = lngResult
    DCencour = False
End Function
