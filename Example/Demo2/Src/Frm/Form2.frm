VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "日志输出"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12855
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   12855
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   12615
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub List1_Click()
    If List1.ListIndex < 0 Then Exit Sub
    Text1.Text = List1.List(List1.ListIndex)
End Sub
