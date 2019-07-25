VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type DllGetClassObjectFun
    rclsid As Long
    riid As Long
    ppv As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByVal lpSource As Long, ByVal cbCopy As Long)
Private Declare Sub StringFromCLSID Lib "ole32.dll" (ByRef rclsid As Any, ByRef lplpsz As Long)
Private WithEvents TEventCollection As SEventCollection
Attribute TEventCollection.VB_VarHelpID = -1
Public Obj As Object
Public Event Events(ByVal Value As SEventInfo)

Public Sub Load(ByVal ComPath As String, ByVal ClsName As String)
    Set Obj = GetObj(ComPath, ClsName)
    Set TEventCollection = New SEventCollection
    TEventCollection.Add Obj, , "SCom"
End Sub

Private Function GetObj(ByVal ComPath As String, ByVal ClsName As String) As Object
    Dim TmpITypeLib As ITypeLib
    Dim TmpCLSIDStr As String
    Dim TmpCLSID As UUID, iid_iunknow As UUID
    Dim TmpClassFactory As IClassFactory
    Dim TmpRetObj As IUnknown
    Dim DllGetClassObjectAddress As Long
    Dim TmpDllGetClassObjectFun As DllGetClassObjectFun
    Set TmpITypeLib = GetComITypeLib(ComPath)
    If TmpITypeLib Is Nothing Then Exit Function
    TmpCLSIDStr = GetComTypeCLSID(TmpCLSID, ClsName, "", TmpITypeLib)
    If TmpCLSIDStr = "" Then Exit Function
    CLSIDFromString IIDSTR_IUnknown, iid_iunknow
    Dim TSDLL As New SDLL
    TSDLL.SDLL_init ComPath
    DllGetClassObjectAddress = TSDLL.SDLL_Bd("DllGetClassObject")
    If DllGetClassObjectAddress = 0 Then Exit Function
    TmpDllGetClassObjectFun.rclsid = VarPtr(TmpCLSID)
    TmpDllGetClassObjectFun.riid = VarPtr(iid_iunknow)
    TmpDllGetClassObjectFun.ppv = VarPtr(TmpClassFactory)
    TSDLL.CallStdcall DllGetClassObjectAddress, VarPtr(TmpDllGetClassObjectFun), 3
    TmpClassFactory.CreateInstance Nothing, iid_iunknow, TmpRetObj
    Set TmpClassFactory = Nothing
    Set GetObj = TmpRetObj
End Function

Private Function GetComITypeLib(ByVal ComPath As String) As ITypeLib
    If Dir(ComPath) = "" Or ComPath = "" Then Exit Function
    Set GetComITypeLib = LoadTypeLibEx(ComPath, REGKIND_NONE)
End Function

Private Function GetComGUID(ByRef OutUUIDType As UUID, Optional ByVal ComPath As String, Optional ByVal TITypeLib As ITypeLib) As String
    If (Dir(ComPath) = "" Or ComPath = "") And (TITypeLib Is Nothing) Then Exit Function
    Dim zc As ITypeLib
    If TITypeLib Is Nothing Then
        Set zc = LoadTypeLibEx(ComPath, REGKIND_NONE)
    Else
        Set zc = TITypeLib
    End If
    Dim ZC2 As Long
    ZC2 = zc.GetLibAttr
    Dim zc3 As TLIBATTR
    CopyMemory zc3, ZC2, LenB(zc3)
    Dim zc4 As Long
    OutUUIDType = zc3.IID
    StringFromCLSID zc3.IID, zc4
    GetComGUID = UBytesToBstr(GetByteFM(zc4, 76))
    zc.ReleaseTLibAttr ZC2
End Function

Private Function GetComTypeCLSID(ByRef OutUUIDType As UUID, ByVal szNameBuf As String, Optional ByVal ComPath As String, Optional ByVal TITypeLib As ITypeLib) As String
    If (Dir(ComPath) = "" Or ComPath = "") And (TITypeLib Is Nothing) Then Exit Function
    Dim zc As ITypeLib
    If TITypeLib Is Nothing Then
        Set zc = LoadTypeLibEx(ComPath, REGKIND_NONE)
    Else
        Set zc = TITypeLib
    End If
    If zc.IsName(szNameBuf, 0) = 0 Then Exit Function
    Dim zzc As ITypeInfo
    Dim zzc2 As Long
    Dim zzc3 As Integer
    zzc3 = 1
    zc.FindName szNameBuf, 0, zzc, zzc2, zzc3
    If zzc3 <> 1 Then Exit Function
    Dim xx2 As Long
    xx2 = zzc.GetTypeAttr
    Dim xxx As TYPEATTR
    CopyMemory xxx, xx2, LenB(xxx)
    Dim xx3 As Long
    OutUUIDType = xxx.IID
    StringFromCLSID xxx.IID, xx3
    GetComTypeCLSID = UBytesToBstr(GetByteFM(xx3, 76))
    zzc.ReleaseTypeAttr xx2
End Function

Private Sub TEventCollection_HandleEvent(ByVal ObjInf As SObjectInfo, ByVal EventInfo As SEventInfo)
    RaiseEvent Events(EventInfo)
End Sub

Private Sub Class_Initialize()
    InitEventSink
End Sub
