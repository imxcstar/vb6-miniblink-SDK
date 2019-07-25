VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SObjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 512
'********************************************************************************
'
' ObjectInfo Class - EventCollection Library
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

Private m_oObject As olelib.IUnknown
Private m_sKey As String
Private m_lCookie As Long
Private m_lIndex As Long
Private m_tIID As UUID

Friend Property Get IID() As UUID
   IID = m_tIID
End Property

Friend Property Get Cookie() As Long
   Cookie = m_lCookie
End Property

Public Property Get Key() As String
   Key = m_sKey
End Property

Public Property Get Object() As Object
   Set Object = m_oObject
End Property

Public Property Get Index() As Long
   Index = m_lIndex
End Property

Friend Sub frInitialize( _
   ByVal Key As String, _
   ByVal Index As Long, _
   ByVal Cookie As Long, _
   ByVal Object As Object, _
   IID As UUID)

   m_sKey = Key
   m_lIndex = Index
   m_lCookie = Cookie
   Set m_oObject = Object
   m_tIID = IID
   
End Sub


