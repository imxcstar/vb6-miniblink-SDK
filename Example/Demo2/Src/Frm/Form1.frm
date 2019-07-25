VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SBrowserTest"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   12585
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   7815
   End
   Begin VB.Menu T1 
      Caption         =   "js回调测试"
   End
   Begin VB.Menu T2 
      Caption         =   "日志输出"
   End
   Begin VB.Menu T3 
      Caption         =   "调试工具"
   End
   Begin VB.Menu T4 
      Caption         =   "测试1"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents mb_callback_com As SCom
Attribute mb_callback_com.VB_VarHelpID = -1
Dim mb_api_com As New SCom
Dim mb_callback As Object
Attribute mb_callback.VB_VarHelpID = -1
Dim mb_api As Object
Dim mb As Long

Private Sub Form_Load()
    Set mb_callback_com = New SCom
    mb_callback_com.Load App.Path & "\MiniblinkSDK.dll", "MiniblinkCallBack"
    mb_api_com.Load App.Path & "\MiniblinkSDK.dll", "MiniblinkAPI"
    
    Set mb_callback = mb_callback_com.Obj
    Set mb_api = mb_api_com.Obj
    
    Me.ScaleMode = 3
    
    mb_api.wkeInitializeEx 0
    
    mb_api.wkeJsBindFunction "test", mb_callback.wkeJsNativeFunction, 0, 2               'js回调事件绑定（影响所有webview和webwindow）
    
    mb = mb_api.wkeCreateWebWindow(2, Me.hWnd, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    mb_api.wkeShowWindow mb, True
    
    mb_api.wkeOnLoadUrlBegin mb, mb_callback.wkeLoadUrlBeginCallback, 0                  'url加载事件绑定
    mb_api.wkeOnCreateView mb, mb_callback.wkeCreateViewCallback, 0                      '创建新窗口事件绑定
    mb_api.wkeOnDownload mb, mb_callback.wkeDownloadCallback, 0                          '下载事件绑定
    mb_api.wkeOnDocumentReady mb, mb_callback.wkeDocumentReadyCallback, 0
    
    mb_api.wkeLoadURL mb, "http://www.baidu.com"
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    Text1.Move 0, 0, Me.ScaleWidth
    mb_api.wkeMoveWindow mb, 0, Text1.Top + Text1.Height, Me.ScaleWidth, Me.ScaleHeight - (Text1.Top + Text1.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form2
    End
End Sub

Private Sub mb_callback_com_Events(ByVal value As SEventInfo)
    'Item数组从1开始
    Dim TmpName As String
    TmpName = value.Name
    Select Case TmpName
    Case "wkeLoadUrlBeginCallback" 'wkeLoadUrlBeginCallback(ByVal webView As Long, ByVal param As Long, ByVal url As String, ByVal job As Long)
        Dim url As String: url = value.Parameters.Item(3)
        AddLog url
    Case "wkeJsNativeFunction" 'wkeJsNativeFunction(ByVal es As Long, ByVal param As Long)
        Dim es As Long: es = value.Parameters.Item(1)
        Dim tret1 As Currency, tret2 As Currency
        tret1 = mb_api.jsArg(es, 0)
        tret2 = mb_api.jsArg(es, 1)
        MsgBox mb_api.jsToTempStringW(es, tret1) & "/" & mb_api.jsToTempStringW(es, tret2)
    Case "wkeDownloadCallback" 'wkeDownloadCallback(ByVal webView As Long, ByVal param As Long, ByVal url As String)
        Dim url2 As String: url2 = value.Parameters.Item(3)
        AddLog "触发了下载事件，下载地址：" & url2
    Case "wkeCreateViewCallback" 'wkeCreateViewCallback(ByVal webView As Long, ByVal param As Long, ByVal navigationType As SBrowser_G.wkeNavigationType, ByVal url As String, windowFeatures As SBrowser_G.wkeWindowFeatures)
        Dim webView As Long: webView = value.Parameters.Item(1)
        AddLog "触发了wkeCreateViewCallback"
        mb_callback.Return_wkeCreateViewCallback = webView      '使用原webview加载
    Case "wkeDocumentReadyCallback" 'wkeDocumentReadyCallback(ByVal webView As Long, ByVal param As Long)
        Dim webView2 As Long: webView2 = value.Parameters.Item(1)
        Text1.Text = mb_api.wkeGetURL(webView2)
    End Select
End Sub

Private Sub T1_Click()
    mb_api.wkeRunJSW mb, "window.test('xcv','hj自行车5gj');"
End Sub

Private Sub AddLog(ByVal value As String)
    If Form2.Visible = False Then Exit Sub
    If Form2.List1.ListCount >= 20000 Then Form2.List1.Clear
    Form2.List1.AddItem value
    Form2.List1.ListIndex = Form2.List1.ListCount - 1
End Sub

Private Sub T2_Click()
    Form2.Show
End Sub

Private Sub T3_Click()
    mb_api.wkeSetDebugConfig mb, "showDevTools", App.Path & "\front_end\inspector.html"
End Sub

Private Sub T4_Click()
    MsgBox mb_api.wkeGetTitle(mb)
End Sub
