Attribute VB_Name = "EventSink"
'********************************************************************************
'
' EventSink Module - EventCollection Library
'
' This module contains the code of the EventSink lightweight object.
'
'********************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on:
'            * Windows XP Pro SP1
'            * VB6 SP5
'
' History:
'           01/02/2003 * This code replaces the old EventCollection
'                        class.
'
'********************************************************************************
Option Explicit

' Event sink object
Type EventSinkData

   ' Pointer to the v-table
   lvtablePtr As Long

   ' Reference count
   RefCount As Long

   ' Interface ID of the
   ' event interface
   EventIID As UUID

   ' Pointer to the owning EventCollection
   EvntColl As Long

   ' Pointer to Object info
   ObjInf As Long

End Type

Private IID_IUnknown As UUID
Private IID_IDispatch As UUID

Private vtable(0 To 6) As Long

' ==== API Declarations ====

Type SAFEARRAY_1D
   cDims As Integer       ' Number of dimensions
   fFeatures As Integer   ' Flags
   cbElements As Long     ' Length of each element
   cLocks As Long         ' Lock count
   pvData As Long         ' Pointer to the data
   Bounds(0 To 0) As SAFEARRAYBOUND   ' Array of dimensions
End Type

Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (PtrDest() As Any) As Long

Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal StrSrc As Long, ByVal StrNew As Long) As Long
Private IsInitEventSink As Boolean

'
' EventSink_QueryInterface
'
' IUnknown::QueryInterface member of the EventSink object
'
Private Function EventSink_QueryInterface(This As EventSinkData, riid As UUID, lObj As Long) As Long

   If IsEqualGUID(riid, IID_IUnknown) Then

      lObj = VarPtr(This)
      This.RefCount = This.RefCount + 1

      #If DEBUG_MODE Then
      Debug.Print "EventSink_QueryInterface(IUknown)"
      #End If

   ElseIf IsEqualGUID(riid, IID_IDispatch) Then

      lObj = VarPtr(This)
      This.RefCount = This.RefCount + 1

      #If DEBUG_MODE Then
      Debug.Print "EventSink_QueryInterface(IDispatch)"
      #End If

   ElseIf IsEqualGUID(riid, This.EventIID) Then

      lObj = VarPtr(This)
      This.RefCount = This.RefCount + 1

      #If DEBUG_MODE Then
      Debug.Print "EventSink_QueryInterface(Event interface)"
      #End If

   Else

      lObj = 0
      EventSink_QueryInterface = E_NOINTERFACE
     
      #If DEBUG_MODE Then
      Debug.Print "EventSink_QueryInterface = E_NOINTERFACE"
      #End If

   End If

End Function

'
' EventSink_AddRef
'
' IUnknown::AddRef member of the EventSink object
'
Private Function EventSink_AddRef(This As EventSinkData) As Long

   This.RefCount = This.RefCount + 1
   EventSink_AddRef = This.RefCount

   #If DEBUG_MODE Then
   Debug.Print "IEventSink.EventSink_AddRef ="; This.RefCount
   #End If

End Function

'
' EventSink_Release
'
' IUnknown::Release member of the EventSink object
'
Private Function EventSink_Release(This As EventSinkData) As Long

   This.RefCount = This.RefCount - 1

   EventSink_Release = This.RefCount

   #If DEBUG_MODE Then
   Debug.Print "IEventSink.EventSink_AddRef = "; This.RefCount
   #End If

   If This.RefCount = 0 Then
      ' Release the object
      GlobalFree VarPtr(This)
   End If

End Function

'
' EventSink_GetTypeInfoCount
'
' IDispatch::GetTypeInfoCount member of the EventSink object
'
Private Function EventSink_GetTypeInfoCount(This As EventSinkData, pctinfo As Long) As Long

   ' Not implemented
   pctinfo = 0
   EventSink_GetTypeInfoCount = E_NOTIMPL

End Function

'
' EventSink_GetTypeInfo
'
' IDispatch::GetTypeInfo member of the EventSink object
'
Private Function EventSink_GetTypeInfo(This As EventSinkData, ByVal iTInfo As Long, ByVal lcid As Long, ppTInfo As Long) As Long

   ' Not implemented
   ppTInfo = 0
   EventSink_GetTypeInfo = E_NOTIMPL

End Function

'
' EventSink_GetIDsOfNames
'
' IDispatch::GetIDsOfNames member of the EventSink object
'
Private Function EventSink_GetIDsOfNames(This As EventSinkData, riid As UUID, rgszNames As Long, ByVal cNames As Long, ByVal lcid As Long, rgDispId As Long) As Long

   ' Not implemented
   EventSink_GetIDsOfNames = E_NOTIMPL

End Function

'
' EventSink_Invoke
'
' IDispatch::Invoke member of the EventSink object
'
Private Function EventSink_Invoke(This As EventSinkData, _
         ByVal dispIdMember As Long, _
         riid As olelib.UUID, _
         ByVal lcid As Long, _
         ByVal wFlags As Integer, _
         ByVal pDispParams As Long, _
         ByVal pVarResult As Long, _
         pExcepInfo As olelib.EXCEPINFO, _
         puArgErr As Long) As Long
Dim oColl As SEventCollection

   On Error Resume Next

   If This.EvntColl <> 0 Then

      #If DEBUG_MODE Then
      Debug.Print "EventSink_Invoke(" & dispIdMember & ")"
      #End If

      ' Get the parent collection
      Set oColl = ResolveObjPtr(This.EvntColl)

      ' Raise the event
      oColl.frRaiseEvent ResolveObjPtr(This.ObjInf), dispIdMember, pDispParams

   End If

   ' This method never fails
   EventSink_Invoke = S_OK

End Function

Private Function AddrOf(ByVal Add As Long) As Long
   AddrOf = Add
End Function

'
' CreateEventSinkObj
'
' Creates a new instance of the EventSink object
'
' Parameters:
' -----------
' EventIID  - IID of the event interface
' ObjInfo   - Reference to the ObjectInfo object
' Coll      - Reference to the EventCollection object
'
Public Function CreateEventSinkObj( _
   EventIID As UUID, _
   ByVal ObjInfo As SObjectInfo, _
   ByVal Coll As SEventCollection) As Object

Dim lEventSinkPtr As Long
Dim lOldProt As Long
Dim EventSink As EventSinkData

   With EventSink

      ' Set the initial reference count to 1
      .RefCount = 1

      ' Save the ID of the events interface
      .EventIID = EventIID

      ' Save a pointer to the parent collection
      MoveMemory .EvntColl, Coll, 4&

      ' Store the object info
      MoveMemory .ObjInf, ObjInfo, 4&

      ' Set the vtable
      .lvtablePtr = VarPtr(vtable(0))
         
   End With
   
   ' Allocate memory for the object
   lEventSinkPtr = GlobalAlloc(GPTR, LenB(EventSink))

   If lEventSinkPtr Then

      ' Copy the structure to the allocated memory
      MoveMemory ByVal lEventSinkPtr, EventSink, LenB(EventSink)

      ' Copy the pointer to the return value
      MoveMemory CreateEventSinkObj, lEventSinkPtr, 4

   Else

      ' Raise the error
      Err.Raise 7, "CreateEventSinkObj"

   End If

End Function

'
' Main
'
' Entry point of the DLL
'
Public Sub InitEventSink()
    If IsInitEventSink = True Then Exit Sub
    IsInitEventSink = True
    
   ' Initialize IIDs
   CLSIDFromString IIDSTR_IUnknown, IID_IUnknown
   CLSIDFromString IIDSTR_IDispatch, IID_IDispatch
   
   ' Initialize EventSink object vtable
   vtable(0) = AddrOf(AddressOf EventSink_QueryInterface)
   vtable(1) = AddrOf(AddressOf EventSink_AddRef)
   vtable(2) = AddrOf(AddressOf EventSink_Release)
   vtable(3) = AddrOf(AddressOf EventSink_GetTypeInfoCount)
   vtable(4) = AddrOf(AddressOf EventSink_GetTypeInfo)
   vtable(5) = AddrOf(AddressOf EventSink_GetIDsOfNames)
   vtable(6) = AddrOf(AddressOf EventSink_Invoke)
   
End Sub

'
' ResolveObjPtr
'
' Returns a strong reference to an object from a pointer
'
Public Function ResolveObjPtr(ByVal Ptr As Long) As olelib.IUnknown
Dim oUnk As olelib.IUnknown
   
   MoveMemory oUnk, Ptr, 4&
   Set ResolveObjPtr = oUnk
   MoveMemory oUnk, 0&, 4&
   
End Function


