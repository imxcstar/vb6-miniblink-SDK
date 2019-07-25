VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SEventCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 256
Attribute VB_Description = "Edanmo's EventCollection Class v2.0"
'********************************************************************************
'
' EventCollection Class
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

' Collection that will hold the objects
Private m_oCollection As Collection

Public UseDISPID As Boolean

'
' HandleEvent
'
' Raised when an event of a contained object is raised
'
Event HandleEvent(ByVal ObjInf As SObjectInfo, ByVal EventInfo As SEventInfo)
'
' pvGetEventName
'
' Returns the name of an event given its DISPID
'
' Parameters:
' -----------
' Source - IDispatch interface of source object
' DISPIP - DISPID of the event
' IID    - Interface ID of the source interface
'
' Return Value: The event name. If the name can't be found
' the function returns the event DISPID.
'
Private Function pvGetEventName( _
      ByVal Source As olelib.IDispatch, _
      ByVal DISPID As Long, _
      IID As olelib.UUID) As Variant
Dim oTL As olelib.ITypeLib
Dim oTI As olelib.ITypeInfo
Dim sEventName As String

   On Error Resume Next

   ' Get the ITypeInfo for the source object
   Set oTI = Source.GetTypeInfo()

   If Not oTI Is Nothing Then

      ' Get the containing type library
      oTI.GetContainingTypeLib oTL

      If Not oTL Is Nothing Then

         ' Find the ITypeInfo of the source interface
         Set oTI = oTL.GetTypeInfoOfIID(IID)

         If Not oTI Is Nothing Then
            oTI.GetNames DISPID, sEventName, 1
         End If

      End If

   End If

   If LenB(sEventName) = 0 Then
      pvGetEventName = DISPID
   Else
      pvGetEventName = sEventName
   End If

End Function
Public Sub Add( _
   Item As Object, _
   Optional SourceIID As String, _
   Optional Key As Variant)
   On Error Resume Next
Dim oObjectInfo As SObjectInfo
Dim oCPC As IConnectionPointContainer
Dim oEnm As IEnumConnectionPoints
Dim oCP As IConnectionPoint
Dim oUnk As olelib.IUnknown
Dim tIID As UUID
Dim lCookie As Long

   ' Get the IConnectionPointContainer interface
   Set oCPC = Item
   
   If LenB(SourceIID) = 0 Then
      
      ' Get connection point enumerator
      Set oEnm = oCPC.EnumConnectionPoints
      
      ' Get the first connection point
      oEnm.Next 1, oCP
      
      ' Get the IID
      oCP.GetConnectionInterface tIID
      
   Else
   
      ' Convert from string to UUID
      CLSIDFromString SourceIID, tIID
      
      ' Get the connection point
      Set oCP = oCPC.FindConnectionPoint(tIID)
      
   End If
      
   ' Create the ObjectInfo object
   Set oObjectInfo = New SObjectInfo
  
   ' Create the event sink object
   Set oUnk = CreateEventSinkObj(tIID, oObjectInfo, Me)
   
   ' Connect the sink object with
   ' the source object
   lCookie = oCP.Advise(oUnk)
   
   On Error GoTo Disconnect
   
   ' Add the object to the collection
   m_oCollection.Add oObjectInfo, Key
   
   ' Initialize the ObjectInfo object
   oObjectInfo.frInitialize Key, _
                            m_oCollection.Count, _
                            lCookie, _
                            Item, _
                            tIID
   
   Exit Sub
   
Disconnect:

   oCP.Unadvise lCookie

   Err.Raise Err.Number, , Err.Description
   
End Sub

Public Sub Clear()
Dim oObjectInfo As SObjectInfo

   ' Disconnect all objects
   For Each oObjectInfo In m_oCollection
      pvUnadvise oObjectInfo
   Next
   
   ' Reset the collection
   Set m_oCollection = New Collection
   
End Sub

Public Property Get Count() As Long
   Count = m_oCollection.Count
End Property

'
' frRaiseEvent
'
' Raises the events. This method is called from each event sink object when a
' event is raised from the source object.
'
' Parameters:
' -----------
' lObjInfo - ObjectInfo object of the source object
' DISPID - DISPID of the event
' lParams - Pointer to a DISPPARAMS structure
'
Friend Sub frRaiseEvent( _
      ByVal ObjInfo As SObjectInfo, _
      ByVal DISPID As Long, _
      ByVal lParams As Long)
Dim oEventInfo As SEventInfo
Dim sEventName As String

   On Error Resume Next
   
   If UseDISPID Then
      sEventName = CStr(DISPID)
   Else
      ' Get the event name
      sEventName = pvGetEventName(ObjInfo.Object, DISPID, ObjInfo.IID)
   End If
   
   ' Create a new EventInfo object
   Set oEventInfo = New SEventInfo
   
   ' Initialize the EventInfo object
   oEventInfo.frInitialize sEventName, pvGetParamArray(lParams)
   
   ' Raise the event
   RaiseEvent HandleEvent(ObjInfo, oEventInfo)
   
   ' Set the event parameters passed ByRef
   pvSetParamArray lParams, oEventInfo.Parameters.frGetParams()

End Sub
Public Property Get Item(Key As Variant) As SObjectInfo

   Set Item = m_oCollection.Item(Key)
   
End Property

Public Function NewEnum() As olelib.IEnumVARIANT
   Set NewEnum = m_oCollection.[_NewEnum]
End Function

'
' pvGetParamArray
'
' Returns the event parameters
'
' Parameters:
' -----------
'
' Ptr - Pointer to a DISPPARAMS structure
'
Private Function pvGetParamArray(ByVal Ptr As Long) As Variant()
Dim tDPAR As olelib.DISPPARAMS
Dim SafeArray As SAFEARRAY_1D
Dim aTmpParams() As Variant
Dim aParams() As Variant
Dim lIdx As Long

   ' Get the DISPARAMS structure form the poitner
   MoveMemory tDPAR, ByVal Ptr, Len(tDPAR)

   ' If there're no parameter exit the function
   If tDPAR.cArgs = 0 Then Exit Function

   ' Fill a SAFEARRAY struct
   ' that points to the parameters
   With SafeArray
      .Bounds(0).cElements = tDPAR.cArgs
      .Bounds(0).lLbound = 0
      .cDims = 1
      .cbElements = 16
      .pvData = tDPAR.rgPointerToVariantArray
   End With

   ' Set the array with the SAFEARRAY struct
   MoveMemory ByVal VarPtrArray(aTmpParams), VarPtr(SafeArray), 4&
   
   ' Copy the array elements
   ReDim aParams(1 To tDPAR.cArgs)
   For lIdx = 1 To tDPAR.cArgs
      aParams(lIdx) = aTmpParams(tDPAR.cArgs - lIdx)
   Next
   
   pvGetParamArray = aParams
   
   ' Release the array
   MoveMemory ByVal VarPtrArray(aTmpParams), 0&, 4&

End Function
'
' pvSetParamArray
'
' Copies the parameters passed ByRef back to the original array
'
' Parameters:
' -----------
'
' Ptr       - Pointer to a DISPPARAMS structure
' aParams() - The parameters array returned by the event handler
'
Private Sub pvSetParamArray(ByVal Ptr As Long, aParams() As Variant)
Dim tDPAR As olelib.DISPPARAMS
Dim SafeArray As SAFEARRAY_1D
Dim iVarType As Integer
Dim lIdx As Long, lPtr As Long
Dim lVarPtr As Long

   ' Get the DISPARAMS structure form the poitner
   MoveMemory tDPAR, ByVal Ptr, Len(tDPAR)
   
   ' If there're no parameter exit the sub
   If tDPAR.cArgs = 0 Then Exit Sub
   
   ' Get the pointer to the first variant
   lVarPtr = tDPAR.rgPointerToVariantArray

   For lIdx = 0 To tDPAR.cArgs - 1
      
      ' Get variant type. Cannot use
      ' VB's VarType because it removes VT_BYREF
      MoveMemory iVarType, ByVal lVarPtr + (lIdx * 16&), 2&
      
      If (iVarType And VT_BYREF) = VT_BYREF Then
         
         ' The parameter was passed ByRef
         
         ' Get the data pointer from the variant
         MoveMemory lPtr, ByVal lVarPtr + (lIdx * 16&) + 8&, 4&
                  
         Select Case iVarType And VT_TYPEMASK
            Case vbString
            
               Dim lStrPtr As Long
               
               ' Get the string pointer
               MoveMemory lStrPtr, ByVal lPtr, 4&
               
               If lStrPtr = 0 Then
               
                  ' Allocate a new string
                  lStrPtr = SysAllocStringPtr(aParams(tDPAR.cArgs - lIdx))
               
               Else
                  
                  ' Reallocate the string
                  lStrPtr = SysReAllocString(lStrPtr, StrPtr(aParams(tDPAR.cArgs - lIdx)))
               
               End If
            
               ' Copy the string pointer
               MoveMemory ByVal lPtr, lStrPtr, 4&
                     
            Case vbInteger
               MoveMemory ByVal lPtr, CInt(aParams(tDPAR.cArgs - lIdx)), 2&
            
            Case vbBoolean
               MoveMemory ByVal lPtr, CBool(aParams(tDPAR.cArgs - lIdx)), 2&
         
            Case vbInteger
               MoveMemory ByVal lPtr, CInt(aParams(tDPAR.cArgs - lIdx)), 2&
         
            Case vbLong
               MoveMemory ByVal lPtr, CLng(aParams(tDPAR.cArgs - lIdx)), 4&
            
            Case vbSingle
               MoveMemory ByVal lPtr, CSng(aParams(tDPAR.cArgs - lIdx)), 4&
            
            Case vbDouble
               MoveMemory ByVal lPtr, CDbl(aParams(tDPAR.cArgs - lIdx)), 8&
            
            Case vbDate
               MoveMemory ByVal lPtr, CDate(aParams(tDPAR.cArgs - lIdx)), 8&
            
            Case vbByte
               MoveMemory ByVal lPtr, CByte(aParams(tDPAR.cArgs - lIdx)), 1&
         
            Case vbCurrency
               MoveMemory ByVal lPtr, CCur(aParams(tDPAR.cArgs - lIdx)), 8&
            
            Case vbVariant, vbDecimal
               VariantCopyIndPtr ByVal lPtr, ByVal VarPtr(aParams(tDPAR.cArgs - lIdx))
            
         End Select
         
      End If
   
   Next
   
End Sub
' pvUnadvise
'
' Disconnects the sink from the connection point
'
Private Sub pvUnadvise(ObjInfo As SObjectInfo)
On Error Resume Next
Dim oCPC As IConnectionPointContainer
Dim oCP As IConnectionPoint

   ' Get the IConnectionPointContainer interface
   Set oCPC = ObjInfo.Object
   
   ' Get the IConnectionPoint
   Set oCP = oCPC.FindConnectionPoint(ObjInfo.IID)
      
   oCP.Unadvise ObjInfo.Cookie
   
End Sub

Public Sub Remove(Key As Variant)
Dim oObjectInfo As SObjectInfo
Dim oCPC As IConnectionPointContainer
Dim oCP As IConnectionPoint

   ' Get the object info
   Set oObjectInfo = m_oCollection(Key)
   
   ' Disconnect the object
   pvUnadvise oObjectInfo
   
   ' Remove the object
   m_oCollection.Remove Key
   
End Sub

Private Sub Class_Initialize()

   ' Initialize the collection
   Set m_oCollection = New Collection
   
End Sub

Private Sub Class_Terminate()
   
   Clear
   Set m_oCollection = Nothing

End Sub




