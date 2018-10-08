VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "搞地址"
   ClientHeight    =   11610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   21675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11610
   ScaleWidth      =   21675
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Caption         =   "贴完企业名称，批量搞一下"
      Height          =   360
      Left            =   15750
      TabIndex        =   8
      Top             =   90
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   11055
      Index           =   2
      Left            =   14580
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   495
      Width           =   7065
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   11055
      Index           =   1
      Left            =   7740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":08CA
      Top             =   495
      Width           =   6795
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   10410
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   7680
      ExtentX         =   13547
      ExtentY         =   18362
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "测试一下"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   50
      Width           =   1485
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   345
      Index           =   0
      Left            =   945
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   11070
      Width           =   6585
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   990
      TabIndex        =   0
      Text            =   "黄山汇金矿业有限公司"
      Top             =   90
      Width           =   3615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "往下面贴一批企业名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8865
      TabIndex        =   9
      Top             =   90
      Width           =   3300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "测试结果："
      Height          =   180
      Left            =   15
      TabIndex        =   4
      Top             =   11160
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "公司名称"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim zzzzzzzzPause  As Boolean

Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox "请输入正确的网址", , "错误！"
        Text1.SetFocus
    Else
        WebBrowser1.Navigate "http://www.qixin.com/search?key=" & UTF8Encode(Text1.Text)
    End If
End Sub
 

Private Sub Command2_Click()
    Dim kk
    Dim i
    kk = Split(Text2(1), vbCrLf)
    If UBound(kk) Then
        For i = 0 To UBound(kk)
            If Len(Trim(kk(i))) Then
            
                Text2(0) = ""
                Text1 = kk(i)
                Call Command1_Click
                
                zzzzzzzzPause = False
                
                Do
                    DoEvents
                Loop While WebBrowser1.ReadyState <> 4
                
                Do
                    DoEvents
                Loop While (zzzzzzzzPause = False)

                If Len(Trim(Text2(0))) Then
                    Text2(2) = Text2(2) & Text2(0) & vbCrLf
                End If
                
            End If
        Next
    End If
End Sub

Private Sub Form_Load()
    zzzzzzzzPause = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub WebBrowser1_DownloadComplete()
   On Error GoTo exitsub
    
    Static lastPos

    Dim WebCont
    Dim pos
    Dim posr
    Err.Clear
    WebCont = WebBrowser1.Document.documentElement.outerHTML
    
    pos = InStr(1, WebCont, "地址：")
    
    If pos Then
        posr = InStr(pos, WebCont, "<")
        Text2(0) = Mid(WebCont, pos + 3, posr - pos - 3)
        If Text2(0) <> lastPos Then
            zzzzzzzzPause = True
            lastPos = Text2(0)
        End If
    End If
exitsub:
    
End Sub

Public Function UTF8Encode(ByVal szInput As String) As String
    Dim wch      As String
    Dim uch      As String
    Dim szRet    As String
    Dim x        As Long
    Dim inputLen As Long
    Dim nAsc     As Long
    Dim nAsc2    As Long
    Dim nAsc3    As Long
     
    If szInput = "" Then
        UTF8Encode = szInput
        Exit Function

    End If

    inputLen = Len(szInput)

    For x = 1 To inputLen
        '得到每个字符
        wch = Mid(szInput, x, 1)
        '得到相应的UNICODE编码
        nAsc = AscW(wch)

        '对于<0的编码　其需要加上65536
        If nAsc < 0 Then nAsc = nAsc + 65536

        '对于<128位的ASCII的编码则无需更改
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else

            If (nAsc And &HF000) = 0 Then
                '真正的第二层编码范围为000080 - 0007FF
                'Unicode在范围D800-DFFF中不存在任何字符，基本多文种平面中约定了这个范围用于UTF-16扩展标识辅助平面（两个UTF-16表示一个辅助平面字符）.
                '当然，任何编码都是可以被转换到这个范围，但在unicode中他们并不代表任何合法的值。
     
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
                 
            Else
                '第三层编码00000800 – 0000FFFF
                '首先取其前四位与11100000进行或去处得到UTF-8编码的前8位
                '其次取其前10位与111111进行并运算，这样就能得到其前10中最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码中间的8位
                '最后将其与111111进行并运算，这样就能得到其最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码最后8位编码
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch

            End If

        End If

    Next
     
    UTF8Encode = szRet

End Function

