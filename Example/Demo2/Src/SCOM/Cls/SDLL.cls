VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal HMODULE As Long, ByVal lpProcName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function CallAsm_4 Lib "user32" Alias "CallWindowProcA" (ByRef lpBytes As Any, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallAsm_1 Lib "user32" Alias "CallWindowProcA" (ByRef lpBytes As Any, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Byte
Private Declare Function CallAsm_8 Lib "user32" Alias "CallWindowProcA" (ByRef lpBytes As Any, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Currency
Private Declare Function CallAsmCode Lib "user32" Alias "CallWindowProcA" (lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Long) As Long
Private h As Long

Public Function SDLL_init(dllname As String) As Long
    h = LoadLibrary(dllname)
    SDLL_init = h
End Function

Public Function SDLL_Bd(funname As String) As Long
    If h = 0 Then SDLL_Bd = 0: Exit Function
    SDLL_Bd = GetProcAddress(h, funname)
End Function

Public Function SDLL_Release() As Long
    If h = 0 Then Exit Function
    SDLL_Release = FreeLibrary(h)
End Function

Public Function CallCdecl_4(lpfn As Long, ParamArray args()) As Long
    Dim ASM() As String
    Dim stub() As Byte
    Dim i As Long
    Dim argSize As Byte
    Dim Ret As Long

    For i = UBound(args) To 0 Step -1
        If Not IsNumeric(args(i)) Then
            Exit Function
        End If
        Push ASM(), "68 " & lng2Hex(CLng(args(i)))
        argSize = argSize + 4
    Next
    
    Push ASM(), "B8 " & lng2Hex(lpfn)
    Push ASM(), "FF D0"
    Push ASM(), "83 C4 " & Hex(argSize)
    Push ASM(), "C2 10 00"
    
    stub() = toBytes(Join(ASM, " "))

    CallCdecl_4 = CallAsm_4(stub(0), 0, 0, 0, 0)
End Function

Public Function CallCdecl_1(lpfn As Long, ParamArray args()) As Byte
    Dim ASM() As String
    Dim stub() As Byte
    Dim i As Long
    Dim argSize As Byte
    Dim Ret As Long

    For i = UBound(args) To 0 Step -1
        If Not IsNumeric(args(i)) Then
            Exit Function
        End If
        Push ASM(), "68 " & lng2Hex(CLng(args(i)))
        argSize = argSize + 4
    Next
    
    Push ASM(), "B8 " & lng2Hex(lpfn)
    Push ASM(), "FF D0"
    Push ASM(), "83 C4 " & Hex(argSize)
    Push ASM(), "C2 10 00"
    
    stub() = toBytes(Join(ASM, " "))

    CallCdecl_1 = CallAsm_1(stub(0), 0, 0, 0, 0)
End Function

Public Function CallCdecl_8(lpfn As Long, ParamArray args()) As Currency
    Dim ASM() As String
    Dim stub() As Byte
    Dim i As Long
    Dim argSize As Byte
    Dim Ret As Long

    For i = UBound(args) To 0 Step -1
        If Not IsNumeric(args(i)) Then
            Exit Function
        End If
        Push ASM(), "68 " & lng2Hex(CLng(args(i)))
        argSize = argSize + 4
    Next
    
    Push ASM(), "B8 " & lng2Hex(lpfn)
    Push ASM(), "FF D0"
    Push ASM(), "83 C4 " & Hex(argSize)
    Push ASM(), "C2 10 00"
    
    stub() = toBytes(Join(ASM, " "))

    CallCdecl_8 = CallAsm_8(stub(0), 0, 0, 0, 0)
End Function

Private Function lng2Hex(x As Long) As String
    Dim b(1 To 4) As Byte
    CopyMemory b(1), x, 4
    lng2Hex = Hex(b(1)) & " " & Hex(b(2)) & " " & Hex(b(3)) & " " & Hex(b(4))
End Function

Private Sub Push(ary, value)
    On Error GoTo init
    Dim x As Long
    x = UBound(ary)
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function toBytes(x As String) As Byte()
    Dim tmp() As String
    Dim fx() As Byte
    Dim i As Long
    
    tmp = Split(x, " ")
    ReDim fx(UBound(tmp))
    
    For i = 0 To UBound(tmp)
        fx(i) = CInt("&h" & tmp(i))
    Next
    
    toBytes = fx()
End Function


'------------------------------------------------------
' 功能：借用API调用任意个数参数的函数的纯VB实现版.
' 此函数内部自带SEH错误处理机制，但并不保证任意
' 错误调用均不会让您的VB崩溃(比如破坏栈平衡后返回)
' 作者：阿国哥　hackor@yeah.net
' 参数一（pFn ）：函数指针
' 参数二（pParam ）：参数指针，指向一个连续的内存块，比如目的函数有三个参数，分别为A，B，C。你可以定义一个结构，结构体为A，B，C（每个参数均为4字节长），然后传这个结构的地址。
' 参数三 (Count): 参数个数?
' 返回: 目的函数的返回值? (你可以修改成其它类型的返回值)
'------------------------------------------------------
Public Function CallStdcall(ByVal pFn As Long, ByVal pParam As Long, ByVal Count As Long) As Long
    Dim CallAnyFuncCode(34) As Long, lRet As Long

    CallAnyFuncCode(0) = &H53EC8B55
    CallAnyFuncCode(1) = &HE8&
    CallAnyFuncCode(2) = &HEB815B00
    CallAnyFuncCode(3) = &H1000112C
    CallAnyFuncCode(4) = &H114A938D
    CallAnyFuncCode(5) = &H64521000
    CallAnyFuncCode(6) = &H35FF&
    CallAnyFuncCode(7) = &H89640000
    CallAnyFuncCode(8) = &H25&
    CallAnyFuncCode(9) = &H8B1FEB00
    CallAnyFuncCode(10) = &HE80C2444
    CallAnyFuncCode(11) = &H0&
    CallAnyFuncCode(12) = &H53E98159
    CallAnyFuncCode(13) = &H8D100011
    CallAnyFuncCode(14) = &H119791
    CallAnyFuncCode(15) = &HB8908910
    CallAnyFuncCode(16) = &H33000000
    CallAnyFuncCode(17) = &H558BC3C0
    CallAnyFuncCode(18) = &H104D8B0C
    CallAnyFuncCode(19) = &HEB8A148D
    CallAnyFuncCode(20) = &HFC528D06
    CallAnyFuncCode(21) = &HB4932FF
    CallAnyFuncCode(22) = &H8BF675C9
    CallAnyFuncCode(23) = &HD0FF0845
    CallAnyFuncCode(24) = &H58F64
    CallAnyFuncCode(25) = &H83000000
    CallAnyFuncCode(26) = &H4D8B04C4
    CallAnyFuncCode(27) = &H5B018914
    CallAnyFuncCode(28) = &H10C2C9
    CallAnyFuncCode(29) = &H58F64
    CallAnyFuncCode(30) = &H83000000
    CallAnyFuncCode(31) = &HC03304C4
    CallAnyFuncCode(32) = &H89144D8B
    CallAnyFuncCode(33) = &HC2C95B21
    CallAnyFuncCode(34) = &H90900010
    CallStdcall = CallAsmCode(CallAnyFuncCode(0), pFn, pParam, Count, lRet)

    If CallStdcall <> lRet Then
        CallStdcall = 0 '这里表示出现严重错误，你应当再了解目的函数的使用方法
        Debug.Assert False '因为你的参数传递问题，导致程序已出现了非法操作。
    End If
End Function





