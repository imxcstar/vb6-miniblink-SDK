VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "�´���"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   10440
   LinkTopic       =   "Form2"
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   696
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu GoHome 
      Caption         =   "�ص���ҳ"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mb_api As New MiniblinkAPI
Public mb As Long

Private Sub GoHome_Click()
    mb_api.wkeLoadURL mb, "http://www.baidu.com"
End Sub
