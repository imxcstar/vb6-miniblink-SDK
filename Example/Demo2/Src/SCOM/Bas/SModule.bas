Attribute VB_Name = "SModule"
Option Explicit
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Function StrTByte(Str As String) As Byte()
    Dim zc() As Byte
    zc = StrConv(Str, vbFromUnicode)
    ReDim Preserve zc(UBound(zc) + 1)
    zc(UBound(zc)) = 0
    StrTByte = zc
End Function

Public Function BytesToBstr(bytes)
    On Error GoTo CuoWu
    Dim SFCW As Boolean
    Dim Unicode As String
    If IsUTF8(bytes) Then
        Unicode = "UTF-8"
    Else
        Unicode = "GB2312"
    End If
TG:
    Dim objstream As Object
    Set objstream = CreateObject("ADODB.Stream")
    With objstream
        .Type = 1
        .Mode = 3
        .Open
        If SFCW = False Then .Write bytes
        .position = 0
        .Type = 2
        .Charset = Unicode
        BytesToBstr = .ReadText
        .Close
    End With
    Exit Function
CuoWu:
    Unicode = "GB2312"
    SFCW = True
    GoTo TG
End Function

Private Function IsUTF8(bytes) As Boolean
    On Error GoTo CuoWu
    Dim i As Long, AscN As Long, length As Long
    length = UBound(bytes) + 1
    
    If length < 3 Then
        IsUTF8 = False
        Exit Function
    ElseIf bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
        IsUTF8 = True
        Exit Function
    End If
    
    Do While i <= length - 1
        If bytes(i) < 128 Then
            i = i + 1
            AscN = AscN + 1
        ElseIf (bytes(i) And &HE0) = &HC0 And (bytes(i + 1) And &HC0) = &H80 Then
            i = i + 2
            
        ElseIf i + 2 < length Then
            If (bytes(i) And &HF0) = &HE0 And (bytes(i + 1) And &HC0) = &H80 And (bytes(i + 2) And &HC0) = &H80 Then
                i = i + 3
            Else
                IsUTF8 = False
                Exit Function
            End If
        Else
            IsUTF8 = False
            Exit Function
        End If
    Loop
    
    If AscN = length Then
        IsUTF8 = False
    Else
        IsUTF8 = True
    End If
    Exit Function
CuoWu:
    IsUTF8 = False
End Function

Public Function pGetStringFromPtr(ByVal lPtr As Long) As String
    Dim Buff() As Byte '声明一个Byte数组
    Dim lPointer As Long '声明一个变量，用于存储指针
    lPointer = lPtr
    ReDim Buff(0 To lstrlen(lPointer) * 2 - 1) As Byte  '分配缓存大小,由于得到的是Unicode，所以乘以2
    lstrcpy Buff(0), ByVal lPointer  '复制到缓存Buff中
    pGetStringFromPtr = BytesToBstr(Buff)
End Function

Public Function UBytesToBstr(bytes() As Byte) As String
    Dim zc() As Byte
    Dim i As Long, j As Long
    For i = 0 To UBound(bytes)
        If bytes(i) <> 0 Then
            ReDim Preserve zc(j) As Byte
            zc(j) = bytes(i)
            j = j + 1
        End If
    Next i
    UBytesToBstr = BytesToBstr(zc)
End Function

Public Function GetByteFM(ByVal address As Long, ByVal length As Long) As Byte()
    Dim zc() As Byte
    ReDim zc(length - 1) As Byte
    CopyMemory zc(0), ByVal address, length
    GetByteFM = zc
End Function

Public Function GetByteFM2(ByVal address As Long) As Byte()
    Dim zc() As Byte
    Dim ZC2 As Long
    ZC2 = 1024
    ReDim zc(ZC2 - 1) As Byte
    CopyMemory zc(0), ByVal address, ZC2
    GetByteFM2 = zc
End Function

