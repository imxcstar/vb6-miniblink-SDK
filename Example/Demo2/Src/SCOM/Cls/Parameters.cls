VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 1024
'********************************************************************************
'
' Parameters Class - EventCollection Library
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

Dim m_avParams() As Variant

Friend Sub frInitialize(Params() As Variant)
   m_avParams = Params
End Sub

Friend Function frGetParams() As Variant()
   frGetParams = m_avParams
End Function

Public Property Get Item(Index As Variant) As Variant
   Item = m_avParams(CLng(Index))
End Property

Public Property Let Item(Index As Variant, Value As Variant)
   m_avParams(CLng(Index)) = Value
End Property

Public Property Set Item(Index As Variant, Value As Variant)
   Set m_avParams(CLng(Index)) = Value
End Property

Public Property Get Count() As Long
   
   On Error GoTo IsEmpty
   
   Count = UBound(m_avParams)
   
   Exit Property
   
IsEmpty:
   
   Count = 0
   
End Property

Public Function NewEnum() As IUnknown
   ' Not yet implemented
End Function

